using System;
using System.Threading;
using System.Windows;

namespace ExcelDna.IntelliSense
{
    // The Popuplist is shown both for function selection, and for some argument selection lists (e.g. TRUE/FALSE).
    // We ignore the reason for showing, and match purely on the text of the selected item.
    class PopupListWatcher : IDisposable
    {
        IntPtr _hwndPopupList = IntPtr.Zero;
        int _selectedItemIndex = -1;

        // NOTE: Event will always be raised on our automation thread
        public event EventHandler SelectedItemChanged;  // Either text or location

        public bool IsVisible{ get; private set; } = false;
        public string SelectedItemText { get; private set; } = string.Empty;
        public Rect SelectedItemBounds { get; private set; } = Rect.Empty;
        public Rect ListBounds { get; private set; } = Rect.Empty;
        public IntPtr PopupListHandle => _hwndPopupList;

        SynchronizationContext _syncContextAuto;
        SynchronizationContext _syncContextMain;
        WindowWatcher _windowWatcher;
        WinEventHook _selectionChangeHook = null;

        public PopupListWatcher(WindowWatcher windowWatcher, SynchronizationContext syncContextAuto, SynchronizationContext syncContextMain)
        {
            _syncContextAuto = syncContextAuto;
            _syncContextMain = syncContextMain;
            _windowWatcher = windowWatcher;
            _windowWatcher.PopupListWindowChanged += _windowWatcher_PopupListWindowChanged;
            _windowWatcher.FormulaEditLocationChanged += _windowWatcher_FormulaEditLocationChanged;
        }

        // Runs on our automation thread
        void _windowWatcher_PopupListWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs e)
        {
            switch (e.Type)
            {
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Create:
                    if (_hwndPopupList == IntPtr.Zero)
                    {
                        Logger.WindowWatcher.Info($"PopupList window created: {e.WindowHandle}");
                    }
                    else
                    {
                        Logger.WindowWatcher.Warn($"PopupList window created more than once!? Old: {_hwndPopupList}, New: {e.WindowHandle}");
                    }

                    _hwndPopupList = e.WindowHandle;
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Destroy:
                    // We expect this only when shutting down
                    Logger.WindowWatcher.Info($"PopupList window destroyed: {e.WindowHandle}");
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Show:
                    Logger.WindowWatcher.Verbose($"PopupList window show");
                    _hwndPopupList = e.WindowHandle;    // We might have missed the Create
                    IsVisible = true;
                    if (_selectionChangeHook == null)
                    {
                        Logger.WindowWatcher.Verbose($"PopupList WinEvent hook initialize");
                        // We set up the structure changed handler so that we can catch the sub-list creation
                        InstallEventHandlers();
                    }
                    UpdateSelectedItem();
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Hide:
                    Logger.WindowWatcher.Verbose($"PopupList window hide");
                    IsVisible = false;
                    UpdateSelectedItem();
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Focus:
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Unfocus:
                    Logger.WindowWatcher.Verbose($"PopupList unexpected focus event!?");
                    break;
                default:
                    break;
            }
        }

        // Runs on our automation thread
        void _windowWatcher_FormulaEditLocationChanged(object sender, EventArgs e)
        {
            if (IsVisible && _selectedItemIndex != -1 && _hwndPopupList != IntPtr.Zero)
            {
                string text;
                Rect itemBounds;
                var hwndListView = Win32Helper.GetFirstChildWindow(_hwndPopupList);
                ListBounds = Win32Helper.GetWindowBounds(_hwndPopupList);
                Win32Helper.GetListViewSelectedItemInfo(hwndListView, out text, out itemBounds);
                itemBounds.Offset(ListBounds.Left, ListBounds.Top);
                SelectedItemBounds = itemBounds;
                OnSelectedItemChanged();
            }
        }
        
        // Runs on our automation thread
        void InstallEventHandlers()
        {
            Logger.WindowWatcher.Verbose($"PopupList Installing event handlers on thread {Thread.CurrentThread.ManagedThreadId}");
            try
            {
                // TODO: Clean up 
                var hwndListView = Win32Helper.GetFirstChildWindow(_hwndPopupList);

                _selectionChangeHook = new WinEventHook(WinEventHook.WinEvent.EVENT_OBJECT_SELECTION, WinEventHook.WinEvent.EVENT_OBJECT_SELECTION, _syncContextAuto, _syncContextMain, hwndListView);
                _selectionChangeHook.WinEventReceived += _selectionChangeHook_WinEventReceived;
                Logger.WindowWatcher.Verbose($"PopupList selection event handler added");
            }
            catch (Exception ex)
            {
                // Probably no longer visible
                Logger.WindowWatcher.Warn($"PopupList event handler error {ex}");
                _hwndPopupList = IntPtr.Zero;
                IsVisible = false;
            }
        }

        void _selectionChangeHook_WinEventReceived(object sender, WinEventHook.WinEventArgs e)
        {
            Logger.WindowWatcher.Verbose($"PopupList PopupListElementSelectedHandler on thread {Thread.CurrentThread.ManagedThreadId}");
            UpdateSelectedItem();
        }

        void UpdateSelectedItem()
        {
            if (_hwndPopupList == IntPtr.Zero)
            {
                Logger.WindowWatcher.Verbose($"PopupList UpdateSelectedItem ignored: PopupList is null");
                return;
            }

            if (!IsVisible)
            {
                if (_selectedItemIndex == -1 &&
                    SelectedItemText == string.Empty &&
                    SelectedItemBounds == Rect.Empty)
                {
                    // Don't change anything, or fire an updated event
                    return;
                }

                // Set to the way things should be when not visible, and fire an updated event
                _selectedItemIndex = -1;
                SelectedItemText = string.Empty;
                SelectedItemBounds = Rect.Empty;
                ListBounds = Rect.Empty;
            }
            else
            {
                string text;
                Rect itemBounds;
                var hwndListView = Win32Helper.GetFirstChildWindow(_hwndPopupList);
                ListBounds = Win32Helper.GetWindowBounds(_hwndPopupList);

                _selectedItemIndex = Win32Helper.GetListViewSelectedItemInfo(hwndListView, out text, out itemBounds);
                if (string.IsNullOrEmpty(text))
                {
                    // We (unexpectedly) failed to get information about the selected item
                    Logger.WindowWatcher.Warn($"PopupList UpdateSelectedItem - IsVisible but GetListViewSelectedItemInfo failed ");
                    _selectedItemIndex = -1;
                    SelectedItemText = string.Empty;
                    SelectedItemBounds = Rect.Empty;
                    ListBounds = Rect.Empty;
                }
                else
                {
                    // Normal case - all is OK
                    itemBounds.Offset(ListBounds.Left, ListBounds.Top);
                    SelectedItemBounds = itemBounds;
                    SelectedItemText = text;
                }
            }
            OnSelectedItemChanged();
        }

        // Raises the event on the automation thread (but the SyncContext.Post here is redundant)
        void OnSelectedItemChanged()
        {
            Logger.WindowWatcher.Verbose($"PopupList SelectedItemChanged {SelectedItemText} ListBounds: {ListBounds}");
            _syncContextAuto.Post(_ => SelectedItemChanged?.Invoke(this, EventArgs.Empty), null);
        }

        public void Dispose()
        {
            Logger.WindowWatcher.Info($"PopupList Dispose Begin");
            _windowWatcher.PopupListWindowChanged -= _windowWatcher_PopupListWindowChanged;
            _windowWatcher = null;
            _selectionChangeHook?.Dispose();
            Logger.WindowWatcher.Info($"PopupList Dispose End");
        }
    }
}
