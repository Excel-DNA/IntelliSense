using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Automation;

namespace ExcelDna.IntelliSense
{
    // The Popuplist is shown both for function selection, and for some argument selection lists (e.g. TRUE/FALSE).
    // We ignore the reason for showing, and match purely on the text of the selected item.
    class PopupListWatcher : IDisposable
    {
        IntPtr            _hwndPopupList;
        AutomationElement _popupList;
        AutomationElement _selectedItem;

        // NOTE: Event will always be raised on our automation thread
        public event EventHandler SelectedItemChanged;  // Either text or location

        public bool IsVisible{ get; set; } = false;
        public string SelectedItemText { get; set; } = string.Empty;
        public Rect SelectedItemBounds { get; set; } = Rect.Empty;
        public IntPtr PopupListHandle => _hwndPopupList;

        SynchronizationContext _syncContextAuto;
        WindowWatcher _windowWatcher;

        public PopupListWatcher(WindowWatcher windowWatcher, SynchronizationContext syncContextAuto)
        {
            _syncContextAuto = syncContextAuto;
            _windowWatcher = windowWatcher;
            _windowWatcher.PopupListWindowChanged += _windowWatcher_PopupListWindowChanged;
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
                    // Automation.AddStructureChangedEventHandler(_popupList, TreeScope.Element, PopupListStructureChangedHandler);
                    // Automation.AddAutomationPropertyChangedEventHandler(_popupList, TreeScope.Element, PopupListVisibleChangedHandler, AutomationElement.???Visible);
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Destroy:
                    // We expect this only when shutting down
                    Logger.WindowWatcher.Info($"PopupList window destroyed: {e.WindowHandle}");
                    try
                    {
                        // DO we need to remove...?
                        if (_popupList != null)
                            Automation.RemoveAutomationPropertyChangedEventHandler(_popupList, PopupListBoundsChanged);
                    }
                    catch (Exception ex)
                    {
                        /// Too Late????
                        Debug.Print($"PopupList Event Handler Remove Error: {ex.Message}");
                    }
                    // Expected when closing
                    // Debug.Assert(false, "PopupList window destroyed...???");
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Show:
                    Logger.WindowWatcher.Verbose($"PopupList window show");
                    IsVisible = true;
                    // We lazy-create the AutomationElement (expecting to make it only once)
                    if (_popupList == null)
                    {
                        Logger.WindowWatcher.Verbose($"PopupList automation initialize");
                        _hwndPopupList = e.WindowHandle;
                        _popupList = AutomationElement.FromHandle(_hwndPopupList);
                        // We set up the structure changed handler so that we can catch the sub-list creation
                        InstallEventHandlers();
                        UpdateSelectedItem();
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Hide:
                    Logger.WindowWatcher.Verbose($"PopupList window hide");
                    IsVisible = false;
                    UpdateSelectedItem(_selectedItem);
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Focus:
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Unfocus:
                    Logger.WindowWatcher.Verbose($"PopupList unexpected focus event!?");
                    break;
                default:
                    break;
            }
        }

        //// Runs on our automation thread
        //void _windowWatcher_MainWindowChanged(object sender, EventArgs args)
        //{
        //    if (_mainWindow != null)
        //    {
        //        Automation.RemoveAutomationPropertyChangedEventHandler(_mainWindow, PopupListBoundsChanged);
        //    }

        //    WindowWatcher windowWatcher = (WindowWatcher)sender;

        //    if (windowWatcher.MainWindow != IntPtr.Zero)
        //    {
        //        // TODO: I've seen an (ElementNotAvailableException) error here that 'the element is not available'.
        //        // TODO: Lots of time-outs here when debugging, but it's probably OK...
        //        try
        //        {
        //            _mainWindow = AutomationElement.FromHandle(windowWatcher.MainWindow);
        //            Automation.AddAutomationPropertyChangedEventHandler(_mainWindow, TreeScope.Element, PopupListBoundsChanged, AutomationElement.BoundingRectangleProperty);
        //        }
        //        catch (Exception ex)
        //        {
        //            Debug.Print($"!!! Error gettting main window from handle: {ex.Message}");
        //        }
        //    }
        //}

        // This runs on an automation event handler thread
        void PopupListBoundsChanged(object sender, AutomationPropertyChangedEventArgs e)
        {
            if (IsVisible && _selectedItem != null)
            {
                // Debug.Print($">>>> PopupListWatcher.LocationChanged on thread {Thread.CurrentThread.ManagedThreadId}");
                UpdateSelectedItem(_selectedItem);
            }

            //_syncContextAuto.Post(delegate
            //{
            //    if (_popupListList != null)
            //    {
            //        //////////////////////
            //        // TODO: What is going on here ...???
            //        //////////////////////
            //        //Automation.AddAutomationEventHandler(
            //        //    SelectionItemPattern.ElementSelectedEvent, _popupListList, TreeScope.Descendants /* was .Children */, PopupListElementSelectedHandler);


            //        var selPat = _popupListList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;

            //        // Update the current selection, if any
            //        var curSel = selPat.Current.GetSelection();
            //        if (curSel.Length > 0)
            //        {
            //            try
            //            {
            //                UpdateSelectedItem(curSel[0]);
            //            }
            //            catch (Exception ex)
            //            {
            //                Debug.Print("Error during UpdateSelected! " + ex);
            //            }
            //        }
            //    }
            //}, null);
        }

        // Runs on an automation event thread
        void PopupListElementSelectedHandler(object sender, AutomationEventArgs e)
        {
            Logger.WindowWatcher.Verbose($"PopupList PopupListElementSelectedHandler on thread {Thread.CurrentThread.ManagedThreadId}");
            UpdateSelectedItem(sender as AutomationElement);
        }

        // Runs on our automation thread
        void InstallEventHandlers()
        {
            Logger.WindowWatcher.Verbose($"PopupList Installing event handlers on thread {Thread.CurrentThread.ManagedThreadId}");
            try
            {
                Automation.AddAutomationEventHandler(
                        SelectionItemPattern.ElementSelectedEvent, _popupList, TreeScope.Descendants /* was .Children */, PopupListElementSelectedHandler);
                Logger.WindowWatcher.Verbose($"PopupList selection event handler added");
                Automation.AddAutomationPropertyChangedEventHandler(_popupList, TreeScope.Element, PopupListBoundsChanged, AutomationElement.BoundingRectangleProperty);
                Logger.WindowWatcher.Verbose($"PopupList bounds change event handler added");
            }
            catch (Exception ex)
            {
                // Probably no longer visible
                Logger.WindowWatcher.Warn($"PopupList event handler error {ex}");
                _popupList = null;
            }
        }

        void UpdateSelectedItem()
        {
            if (_popupList == null)
            {
                Logger.WindowWatcher.Verbose($"PopupList UpdateSelectedItem ignored: PopupList is null");
                return;
            }

            Condition patCondition = new PropertyCondition(
                AutomationElement.IsSelectionPatternAvailableProperty, true);
            var listElement = _popupList.FindFirst(TreeScope.Descendants, patCondition);
            if (listElement != null)
            {
                var selectionPattern = listElement.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;
                var currentSelection = selectionPattern.Current.GetSelection();
                if (currentSelection.Length > 0)
                {
                    try
                    {
                        UpdateSelectedItem(currentSelection[0]);
                    }
                    catch (Exception ex)
                    {
                        Logger.WindowWatcher.Warn($"PopupList UpdateSelectedItem error {ex}");
                    }
                }
            }
            else
            {
                Logger.WindowWatcher.Warn("PopupList UpdateSelectedItem - No descendent has SelectionPattern !?");
            }
        }

        //// TODO: This should be exposed as an event and popup resize should be elsewhere
        //// Runs on an automation event thread
        //void PopupListStructureChangedHandler(object sender, StructureChangedEventArgs e)
        //{
        //    Debug.Print($">>>> PopupListWatcher.PopupListStructureChangedHandler ({e.StructureChangeType}) on thread {Thread.CurrentThread.ManagedThreadId}");
        //    // Debug.WriteLine($">>> PopupList structure changed - {e.StructureChangeType}");
        //    // CONSIDER: Others too?
        //    if (e.StructureChangeType == StructureChangeType.ChildAdded)
        //    {
        //        var functionList = sender as AutomationElement;

        //        //var listRect = (Rect)functionList.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);

        //        var listElement = functionList.FindFirst(TreeScope.Children, Condition.TrueCondition);
        //        if (listElement != null)
        //        {
        //            // Debug.Print($">>>> PopupListWatcher.PopupListStructureChangedHandler Post - Children Found !!!");

        //            _popupListList = listElement;

        //            // TODO: Move this code!
        //            // TestMoveWindow(_popupListList, (int)listRect.X, (int)listRect.Y);
        //            // TestMoveWindow(functionList, 0, 0);

        //            // If the _popupListList automation element is no plonger valid then we seem to get a InvalidPattern exception here.
        //            var selPat = _popupListList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;
        //            Debug.Assert(selPat != null);

        //            // CONSIDER: Send might be a bit disruptive here / might not be necessary...
        //            // _syncContextAuto.Send( _ =>
        //            //{
        //                try
        //                {
        //                    Debug.Print("POPUPLISTWATCHER WINDOW SELECTION EVENT HANDLER ADDED");
        //                    Automation.AddAutomationEventHandler(
        //                        SelectionItemPattern.ElementSelectedEvent, _popupListList, TreeScope.Descendants /* was .Children */, PopupListElementSelectedHandler);
        //                }
        //                catch (Exception ex)
        //                {
        //                    Debug.Print("Error during AddAutomationEventHandler! " + ex);
        //                }
        //            //}, null);

        //            // Update the current selection, if any
        //            var curSel = selPat.Current.GetSelection();
        //            if (curSel.Length > 0)
        //            {
        //                try
        //                {
        //                    UpdateSelectedItem(curSel[0]);
        //                }
        //                catch (Exception ex)
        //                {
        //                    Debug.Print("Error during UpdateSelected! " + ex);
        //                }
        //            }
        //        }
        //        else
        //        {
        //            Debug.Print($">>>> PopupListWatcher.PopupListStructureChangedHandler Post - No Children Found ??? ");
        //            Debug.WriteLine("ERROR!!! Structure changed but no children anymore.");
        //        }
        //    }
        //    else if (e.StructureChangeType == StructureChangeType.ChildRemoved)
        //    {
        //        if (_popupListList != null)
        //        {
        //            var selPat = _popupListList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;
        //            Debug.Assert(selPat != null);
        //            // CONSIDER: Send might be a bit disruptive here / might not be necessary...
        //            //_syncContextAuto.Send( _ =>
        //            //{
        //                try
        //                {
        //                    Debug.Print("POPUPLISTWATCHER WINDOW SELECTION EVENT HANDLER REMOVED");
        //                    Automation.RemoveAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent, _popupListList, PopupListElementSelectedHandler);
        //                }
        //                catch (Exception ex)
        //                {
        //                    Debug.Print("Error during RemoveAutomationEventHandler! " + ex);
        //                }
        //            //}, null);
        //            _popupListList = null;
        //        }
        //        _selectedItem = null;
        //        OnSelectedItemChanged();
        //    }
        //}

        // Can run on our automation thread or on any automation event thread (which is also allowed to read properties)
        // But might fail, if the newSelectedItem is already gone by the time we run...
        void UpdateSelectedItem(AutomationElement newSelectedItem)
        {
            // Debug.Print($"POPUPLISTWATCHER WINDOW CURRENT SELECTION {newSelectedItem}");

            // TODO: Sometimes the IsVisble is not updated, but we are visible and the first selection is set

            if (!IsVisible || newSelectedItem == null)
            {
                if (_selectedItem == null &&
                    SelectedItemText == string.Empty &&
                    SelectedItemBounds == Rect.Empty)
                {
                    // Don't change and fire event
                    return;
                }
                _selectedItem = null;
                SelectedItemText = string.Empty;
                SelectedItemBounds = Rect.Empty;
            }
            else
            {
                string selectedItemText = string.Empty;
                Rect selectedItemBounds = Rect.Empty;

                try
                {
                    selectedItemText = (string)newSelectedItem.GetCurrentPropertyValue(AutomationElement.NameProperty);
                    selectedItemBounds = (Rect)newSelectedItem.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
                }
                catch (Exception ex)
                {
                    Logger.WindowWatcher.Warn($"PopupList - Could not update selected item: {ex}");
                    // Don't fire the event - we couldn't process this change
                    return;
                }

                _selectedItem = newSelectedItem;
                SelectedItemText = selectedItemText;
                SelectedItemBounds = selectedItemBounds;
                // Debug.Print($"SelectedItemBounds: {SelectedItemBounds}");
            }
            OnSelectedItemChanged();
        }

        // Raises the event on the automation thread (but the SyncContext.Post here is redundant)
        void OnSelectedItemChanged()
        {
            Logger.WindowWatcher.Verbose($"PopupList SelectedItemChanged {SelectedItemText}");
            _syncContextAuto.Post(_ => SelectedItemChanged(this, EventArgs.Empty), null);
        }

        public void Dispose()
        {
            Logger.WindowWatcher.Info($"PopupList Dispose Begin");
            _windowWatcher.PopupListWindowChanged -= _windowWatcher_PopupListWindowChanged;
            _windowWatcher = null;

            _syncContextAuto.Send(delegate
            {
                Debug.Print("Disposing PopupListWatcher - In Automation context");
                if (_popupList != null)
                {
                    //Automation.RemoveAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent, _popupList, PopupListElementSelectedHandler);
                    //Automation.RemoveAutomationPropertyChangedEventHandler(_popupList, PopupListBoundsChanged);
                    _popupList = null;
                }
            }, null);
            Logger.WindowWatcher.Info($"PopupList Dispose End");
        }
    }
}
