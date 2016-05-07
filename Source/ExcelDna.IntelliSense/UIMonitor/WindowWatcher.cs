using System;
using System.Diagnostics;
using System.Threading;
using System.Windows.Automation;

namespace ExcelDna.IntelliSense
{
    // NOTE: We're using the COM wrapper (UIAComWrapper) approach to UI Automation, rather than using the old System.Automation in .NET BCL
    //       My understanding is from this post : https://social.msdn.microsoft.com/Forums/en-US/69cf1072-d57f-4aa1-a8ea-ea8a9a5da70a/using-uiautomation-via-com-interopuiautomationclientdll-causes-windows-explorer-to-crash?forum=windowsaccessibilityandautomation
    //       I think this is a sample of the new UI Automation 3.0 COM API: http://blogs.msdn.com/b/winuiautomation/archive/2012/03/06/windows-7-ui-automation-client-api-c-sample-e-mail-reader-version-1-0.aspx
    //       Really need to keep in mind the threading guidance from here: https://msdn.microsoft.com/en-us/library/windows/desktop/ee671692(v=vs.85).aspx

    // NOTE: Check that the TextPattern is available where expected under Windows 7 (remember about CLSID_CUIAutomation8 class)

    // NOTE: More ideas on tracking the current topmost window (using Application events): http://www.jkp-ads.com/Articles/keepuserformontop02.asp

    // All the code in this class runs on the Automation thread, including events we handle from the WinEventHook.
    class WindowWatcher : IDisposable
    {
        public class WindowChangedEventArgs : EventArgs
        {

            // CONSIDER: This mapping is trivial - just get rid of it?
            public enum ChangeType
            {
                Create = 1,
                Destroy = 2,
                Show = 3,
                Hide = 4,
                Focus = 5
            }

            public readonly IntPtr WindowHandle;
            public readonly ChangeType Type;

            internal WindowChangedEventArgs(IntPtr windowHandle, WinEventHook.WinEvent winEvent)
            {
                WindowHandle = windowHandle;
                switch (winEvent)
                {
                    case WinEventHook.WinEvent.EVENT_OBJECT_CREATE:
                        Type = ChangeType.Create;
                        break;
                    case WinEventHook.WinEvent.EVENT_OBJECT_DESTROY:
                        Type = ChangeType.Destroy;
                        break;
                    case WinEventHook.WinEvent.EVENT_OBJECT_SHOW:
                        Type = ChangeType.Show;
                        break;
                    case WinEventHook.WinEvent.EVENT_OBJECT_HIDE:
                        Type = ChangeType.Hide;
                        break;
                    case WinEventHook.WinEvent.EVENT_OBJECT_FOCUS:
                        Type = ChangeType.Focus;
                        break;
                    default:
                        throw new ArgumentException("Unexpected WinEvent type", nameof(winEvent));
                }
            }

            internal static bool IsSupportedWinEvent(WinEventHook.WinEvent winEvent)
            {
                return winEvent == WinEventHook.WinEvent.EVENT_OBJECT_CREATE ||
                       winEvent == WinEventHook.WinEvent.EVENT_OBJECT_DESTROY ||
                       winEvent == WinEventHook.WinEvent.EVENT_OBJECT_SHOW ||
                       winEvent == WinEventHook.WinEvent.EVENT_OBJECT_HIDE ||
                       winEvent == WinEventHook.WinEvent.EVENT_OBJECT_FOCUS;
            }
        }
        
        const string _mainWindowClass = "XLMAIN";
        const string _sheetWindowClass = "EXCEL7";  // This is the sheet portion (not top level) - we get some notifications from here?
        const string _formulaBarClass = "EXCEL<";
        const string _inCellEditClass = "EXCEL6";
        const string _popupListClass = "__XLACOOUTER";
        const string _nuiDialogClass = "NUIDialog";
        const string _selectDataSourceTitle = "Select Data Source";     // TODO: How does localization work?

        WinEventHook _windowStateChangeHook;


//        public IntPtr SelectDataSourceWindow { get; private set; }
//        public bool IsSelectDataSourceWindowVisible { get; private set; }

        // NOTE: The WindowWatcher raises all events on our Automation thread (via syncContextAuto passed into the constructor).
        // Raised for every WinEvent related to window of the relevant class
        public event EventHandler<WindowChangedEventArgs> FormulaBarWindowChanged;
        public event EventHandler<WindowChangedEventArgs> InCellEditWindowChanged;
        public event EventHandler<WindowChangedEventArgs> PopupListWindowChanged;
//        public event EventHandler<WindowChangedEventArgs> SelectDataSourceWindowChanged;

        public WindowWatcher(SynchronizationContext syncContextAuto)
        {
            #pragma warning disable CS0618 // Type or member is obsolete (GetCurrentThreadId) - But for debugging we want to monitor this anyway
            // Debug.Print($"### WindowWatcher created on thread: Managed {Thread.CurrentThread.ManagedThreadId}, Native {AppDomain.GetCurrentThreadId()}");
            #pragma warning restore CS0618 // Type or member is obsolete

            // Using WinEvents instead of Automation so that we can watch top-level window changes, but only from the right (current Excel) process.
            // TODO: We need to dramatically reduce the number of events we grab here...
            _windowStateChangeHook = new WinEventHook(WinEventHook.WinEvent.EVENT_OBJECT_CREATE, WinEventHook.WinEvent.EVENT_OBJECT_FOCUS, syncContextAuto);
//            _windowStateChangeHook = new WinEventHook(WinEventHook.WinEvent.EVENT_OBJECT_CREATE, WinEventHook.WinEvent.EVENT_OBJECT_END, syncContextAuto);

            _windowStateChangeHook.WinEventReceived += _windowStateChangeHook_WinEventReceived;
        }

        // Runs on the Automation thread (before syncContextAuto starts pumping)
        public void TryInitialize()
        {
            // Debug.Print("### WindowWatcher TryInitialize on thread: " + Thread.CurrentThread.ManagedThreadId);
            AutomationElement focused;
            try
            {
                focused = AutomationElement.FocusedElement;
                var className = focused.GetCurrentPropertyValue(AutomationElement.ClassNameProperty);
                Debug.Print($"WindowWatcher.TryInitialize - Focused: {className}");
            }
            catch (ArgumentException aex)
            {
                Debug.Print($"!!! ERROR: Failed to get Focused Element: {aex}");
                return;
            }
        }

        // This runs on the Automation thread, via SyncContextAuto passed in to WinEventHook when we created this WindowWatcher
        // CONSIDER: We would be able to run all the watcher updates from WinEvents, including Location and Selection changes,
        //           but since WinEvents have no hwnd filter, UIAutomation events might be more efficient.
        // CONSIDER: Performance optimisation would keep a list of window handles we know about, preventing the class name check every time
        void _windowStateChangeHook_WinEventReceived(object sender, WinEventHook.WinEventArgs e)
        {
            if (!WindowChangedEventArgs.IsSupportedWinEvent(e.EventType))
                return;

            // Debug.Print("### Thread receiving WindowStateChange: " + Thread.CurrentThread.ManagedThreadId);
            var className = Win32Helper.GetClassName(e.WindowHandle);
            switch (className)
            {
                //case _sheetWindowClass:
                //    if (e.EventType == WinEventHook.WinEvent.EVENT_OBJECT_SHOW)
                //    {
                //        // Maybe a new workbook is on top...
                //        // Note that there is also an EVENT_OBJECT_PARENTCHANGE (which we are not subscribing to at the moment
                //    }
                //    break;
                case _popupListClass:
                    PopupListWindowChanged?.Invoke(this, new WindowChangedEventArgs(e.WindowHandle, e.EventType));
                    break;
                case _inCellEditClass:
                    InCellEditWindowChanged?.Invoke(this, new WindowChangedEventArgs(e.WindowHandle, e.EventType));
                    break;
                case _formulaBarClass:
                    FormulaBarWindowChanged?.Invoke(this, new WindowChangedEventArgs(e.WindowHandle, e.EventType));
                    break;
                //case _nuiDialogClass:
                //    // Debug.Print($"SelectDataSource {_selectDataSourceClass} Window update: {e.WindowHandle:X}, EventType: {e.EventType}, idChild: {e.ChildId}");
                //    if (e.EventType == WinEventHook.WinEvent.EVENT_OBJECT_CREATE)
                //    {
                //        // Get the name of this window - maybe ours or some other NUIDialog
                //        var windowTitle = Win32Helper.GetText(e.WindowHandle);
                //        if (windowTitle.Equals(_selectDataSourceTitle, StringComparison.OrdinalIgnoreCase))
                //        {
                //            SelectDataSourceWindow = e.WindowHandle;
                //            SelectDataSourceWindowChanged?.Invoke(this, 
                //                new WindowChangedEventArgs { Type = WindowChangedEventArgs.ChangeType.Create });
                //        }
                //    }
                //    else if (SelectDataSourceWindow == e.WindowHandle && e.EventType == WinEventHook.WinEvent.EVENT_OBJECT_SHOW)
                //    {
                //        IsSelectDataSourceWindowVisible = true;
                //        SelectDataSourceWindowChanged?.Invoke(this,
                //                new WindowChangedEventArgs { Type = WindowChangedEventArgs.ChangeType.Create });
                //    }
                //    else if (SelectDataSourceWindow == e.WindowHandle && e.EventType == WinEventHook.WinEvent.EVENT_OBJECT_HIDE)
                //    {
                //        IsSelectDataSourceWindowVisible = false;
                //        SelectDataSourceWindowChanged?.Invoke(this, new WindowChangedEventArgs { Type = WindowChangedEventArgs.ChangeType.Hide });
                //    }
                //    else if (SelectDataSourceWindow == e.WindowHandle && e.EventType == WinEventHook.WinEvent.EVENT_OBJECT_DESTROY)
                //    {
                //        IsSelectDataSourceWindowVisible = false;
                //        SelectDataSourceWindow = IntPtr.Zero;
                //        SelectDataSourceWindowChanged?.Invoke(this, new WindowChangedEventArgs { Type = WindowChangedEventArgs.ChangeType.Destroy });
                //    }
                //    break;
                default:
                    if (e.EventType == WinEventHook.WinEvent.EVENT_OBJECT_FOCUS)
                    {
                        Logger.WindowWatcher.Verbose($"FOCUS! on {className}");
                    }
                    //InCellEditWindowChanged(this, EventArgs.Empty);
                    break;
            }
        }

        public void Dispose()
        {
            if (_windowStateChangeHook != null)
            {
                _windowStateChangeHook.Dispose();
                _windowStateChangeHook = null;
            }
        }
    }


 

    //class SelectDataSourceWatcher : IDisposable
    //{
    //    SynchronizationContext _syncContextAuto;
    //    WindowWatcher _windowWatcher;

    //    public IntPtr SelectDataSourceWindow { get; private set; }
    //    public event EventHandler SelectDataSourceWindowChanged;
    //    public bool IsVisible { get; private set; }
    //    public event EventHandler IsVisibleChanged;

    //    public SelectDataSourceWatcher(WindowWatcher windowWatcher, SynchronizationContext syncContextAuto)
    //    {
    //        _syncContextAuto = syncContextAuto;
    //        _windowWatcher = windowWatcher;
    //        _windowWatcher.SelectDataSourceWindowChanged += _windowWatcher_SelectDataSourceWindowChanged;
    //        //_windowWatcher.MainWindowChanged += _windowWatcher_MainWindowChanged;
    //        //_windowWatcher.PopupListWindowChanged += _windowWatcher_PopupListWindowChanged;
    //    }

    //    void _windowWatcher_SelectDataSourceWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs e)
    //    {
    //    }

    //    public void Dispose()
    //    {
    //        Debug.Print("Disposing SelectDataSourceWatcher");
    //        // CONSIDER: Do we need this?
    //        //_windowWatcher.MainWindowChanged -= _windowWatcher_MainWindowChanged;
    //        //_windowWatcher.PopupListWindowChanged -= _windowWatcher_PopupListWindowChanged;
    //        _windowWatcher.SelectDataSourceWindowChanged -= _windowWatcher_SelectDataSourceWindowChanged;
    //        _windowWatcher = null;

    //        _syncContextAuto.Post(delegate
    //        {
    //            Debug.Print("Disposing SelectDataSourceWatcher - In Automation context");
    //        }, null);
    //    }
    //}
}
