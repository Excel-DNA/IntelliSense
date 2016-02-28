using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using Point = System.Drawing.Point;

namespace ExcelDna.IntelliSense
{
    // NOTE: Really need to understand the two approaches to using UI Automation (BCL classes vs. COM wrapper)
    //       My understanding from this post : https://social.msdn.microsoft.com/Forums/en-US/69cf1072-d57f-4aa1-a8ea-ea8a9a5da70a/using-uiautomation-via-com-interopuiautomationclientdll-causes-windows-explorer-to-crash?forum=windowsaccessibilityandautomation
    //       is that we might prefer to use the UIAComWrapper.
    //       On the other hand, "Even without UiaComWrapper, you can always use UI Automation in .NET through COM Interop -- just add a reference to the UIAutomationCore.dll library for your project. "
    //       I think this is a sample of the new UI Automation 3.0 COM API: http://blogs.msdn.com/b/winuiautomation/archive/2012/03/06/windows-7-ui-automation-client-api-c-sample-e-mail-reader-version-1-0.aspx
    //       Really need to understand threading guidance from here: https://msdn.microsoft.com/en-us/library/windows/desktop/ee671692(v=vs.85).aspx

    class WindowWatcher : IDisposable
    {
        WinEventHook _windowStateChangeHook;

        public IntPtr MainWindow { get; private set; }
        public IntPtr FormulaBarWindow { get; private set; }
        public IntPtr InCellEditWindow { get; private set; }
        public IntPtr PopupListWindow { get; private set; }
        public IntPtr PopupListList { get; private set; }
        public IntPtr SelectDataSourceWindow { get; private set; }

        // CONSIDER: The WindowWatcher might rather raise all events on the UIAutomation thread.
        public event EventHandler FormulaBarWindowChanged = delegate { };
        public event EventHandler FormulaBarFocused = delegate { };
        public event EventHandler InCellEditWindowChanged = delegate { };
        public event EventHandler InCellEditFocused = delegate { };
        public event EventHandler MainWindowChanged = delegate { };
        public event EventHandler PopupListWindowChanged = delegate { };   // Might start off with nothing. Changes at most once.
        public event EventHandler PopupListListChanged = delegate { };   // Might start off with nothing. Changes at most once.
        public event EventHandler SelectDataSourceWindowChanged = delegate { };   // Might start off with nothing. Changes at most once.

        public WindowWatcher(string xllName)
        {
            Debug.Print("### WindowWatcher created on thread: " + Thread.CurrentThread.ManagedThreadId);

            // Using WinEvents instead of Automation so that we can watch top-level changes, but only from the right process.

            _windowStateChangeHook = new WinEventHook(WindowStateChange,
                WinEventHook.WinEvent.EVENT_OBJECT_CREATE, WinEventHook.WinEvent.EVENT_OBJECT_FOCUS, xllName);
        }

        public void TryInitialize()
        {
            Debug.Print("### WindowWatcher TryInitialize on thread: " + Thread.CurrentThread.ManagedThreadId);

            AutomationElement focused;
            try
            {
                focused = AutomationElement.FocusedElement;
            }
            catch (ArgumentException aex)
            {
                Debug.Print("!!! ERROR: Failed to get Focused Element: " + aex.ToString());
                return;
            }
            AutomationElement current = focused;

            TreeWalker treeWalker = TreeWalker.ControlViewWalker;
            while ((string)current.GetCurrentPropertyValue(AutomationElement.ClassNameProperty) != _classMain)
            {
                Debug.Print("Current Class = " + (string)current.GetCurrentPropertyValue(AutomationElement.ClassNameProperty));
                current = treeWalker.GetParent(current);
                if (current == null) return;    // At the root
            }
            UpdateMainWindow((IntPtr)(int)current.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty));
        }

        void WindowStateChange(IntPtr hWinEventHook, WinEventHook.WinEvent eventType, IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
        {
            // This runs on the main application thread.
            // We have to get off this thread very quickly.
            var className = Win32Helper.GetClassName(hWnd);
            Logger.WinEvents.Verbose($"{eventType} - hWnd: {hWnd:X} ({className}) - IDs: {idObject} / {idChild} - Thread: {dwEventThread}");
            Debug.Print("### Thread receiving WindowStateChange: " + Thread.CurrentThread.ManagedThreadId);
            switch (className)
            {
                case _classMain:
                    //if (eventType == WinEventHook.WinEvent.EVENT_OBJECT_FOCUS ||
                    //    eventType == WinEventHook.WinEvent.EVENT_OBJECT_SHOW)
                    {
                        Debug.Print("MainWindow update: " + hWnd.ToString("X") + ", EventType: " + eventType);
                        UpdateMainWindow(hWnd);
                    }
                    if (eventType == WinEventHook.WinEvent.EVENT_OBJECT_DESTROY)
                    {

                    }
                    break;
                case _classPopupList:
                    if (PopupListWindow == IntPtr.Zero &&
                        (eventType == WinEventHook.WinEvent.EVENT_OBJECT_CREATE ||
                         eventType == WinEventHook.WinEvent.EVENT_OBJECT_SHOW))
                    {
                        PopupListWindow = hWnd;
                        PopupListWindowChanged(this, EventArgs.Empty);
                    }
                    else
                    {
                        Debug.Assert(PopupListWindow == hWnd);
                    }
                    break;
                case _classInCellEdit:

                    Debug.Print("InCell Window update: " + hWnd.ToString("X") + ", EventType: " + eventType + ", idChild: " + idChild);
                    if (eventType == WinEventHook.WinEvent.EVENT_OBJECT_CREATE ||
                         eventType == WinEventHook.WinEvent.EVENT_OBJECT_SHOW)
                    {
                        InCellEditWindow = hWnd;
                        InCellEditWindowChanged(this, EventArgs.Empty);
                    }
                    else if (eventType == WinEventHook.WinEvent.EVENT_OBJECT_HIDE)
                    {
                        InCellEditWindow = IntPtr.Zero;
                        InCellEditWindowChanged(this, EventArgs.Empty);
                    }
                    else if (eventType == WinEventHook.WinEvent.EVENT_OBJECT_FOCUS)
                    {
                        InCellEditFocused(this, EventArgs.Empty);
                    }
                    break;
                case _classFormulaBar:
                    Debug.Print("FormulaBar Window update: " + hWnd.ToString("X") + ", EventType: " + eventType + ", idChild: " + idChild);
                    if (eventType == WinEventHook.WinEvent.EVENT_OBJECT_CREATE ||
                         eventType == WinEventHook.WinEvent.EVENT_OBJECT_SHOW)
                    {
                        FormulaBarWindow = hWnd;
                        FormulaBarWindowChanged(this, EventArgs.Empty);
                    }
                    else if (eventType == WinEventHook.WinEvent.EVENT_OBJECT_FOCUS)
                    {
                        FormulaBarFocused(this, EventArgs.Empty);
                    }
                    break;
                case _classSelectSeriesData:
                    Debug.Print("SelectSeriesData Window update: " + hWnd.ToString("X") + ", EventType: " + eventType + ", idChild: " + idChild);

                    break;
                default:
                    //InCellEditWindowChanged(this, EventArgs.Empty);
                    break;
            }
        }

        private void UpdateMainWindow(IntPtr hWnd)
        {
            if (MainWindow != hWnd)
            {
                MainWindow = hWnd;
                FormulaBarWindow = IntPtr.Zero;
                MainWindowChanged(this, EventArgs.Empty);
            }

            if (FormulaBarWindow == IntPtr.Zero)
            {
                // Either fresh window, or children not yet set up

                // Search for formulaBar
                AutomationElement mainWindow = AutomationElement.FromHandle(hWnd);
#if DEBUG
                var mainChildren = mainWindow.FindAll(TreeScope.Children, Condition.TrueCondition);
                foreach (AutomationElement child in mainChildren)
                {
                    var hWndChild = (IntPtr)(int)child.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);
                    var classChild = (string)child.GetCurrentPropertyValue(AutomationElement.ClassNameProperty);
                    Debug.Print("Child: " + hWndChild + ", Class: " + classChild);
                }
#endif
                AutomationElement formulaBar = mainWindow.FindFirst(TreeScope.Children,
                    new PropertyCondition(AutomationElement.ClassNameProperty, _classFormulaBar));
                if (formulaBar != null)
                {
                    FormulaBarWindow = (IntPtr)(int)formulaBar.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);

                    // CONSIDER:
                    //  Watch WindowClose event for MainWindow?

                    FormulaBarWindowChanged(this, EventArgs.Empty);
                }
                else
                {
                    Debug.Print("Could not get FormulaBar!");
                }
            }
        }

        const string _classMain = "XLMAIN";
        const string _classFormulaBar = "EXCEL<";
        const string _classInCellEdit = "EXCEL6";
        const string _classPopupList = "__XLACOOUTER";
        const string _classSelectSeriesData = "NUIDialog";

        public void Dispose()
        {
            if (_windowStateChangeHook != null)
            {
                _windowStateChangeHook.Stop();
                _windowStateChangeHook = null;
            }
        }
    }

    enum StateChangeTypeEnum
    {
        Undefined,
        Move
    }

    class StateChangeEventArgs : EventArgs
    {
        public StateChangeTypeEnum StateChangeType { get; private set; }

        public StateChangeEventArgs(StateChangeTypeEnum? stateChangeType = null)
        {
            StateChangeType = stateChangeType ?? StateChangeTypeEnum.Undefined;
        }
    }

    // We want to know whether to show the function entry help
    // For now we ignore whether another ToolTip is being shown, and just use the formula edit state.
    class FormulaEditWatcher : IDisposable
    {
        AutomationElement _formulaBar;
        AutomationElement _inCellEdit;
        AutomationElement _mainWindow;
         
        public event EventHandler<StateChangeEventArgs> StateChanged = delegate { };

        public bool IsEditingFormula { get; set; }
        //        public string CurrentFormula { get; set; } // Easy to get, but we don't need it
        public string CurrentPrefix { get; set; }    // Null if not editing
        // We don't really care whether it is the formula bar or in-cell, 
        // we just need to get the right window handle 
        public Rect EditWindowBounds { get; set; }
        public Point CaretPosition { get; set; }

        private SynchronizationContext _syncContext;

        public FormulaEditWatcher(WindowWatcher windowWatcher, SynchronizationContext syncContext)
        {
            windowWatcher.FormulaBarWindowChanged += delegate { WatchFormulaBar(windowWatcher.FormulaBarWindow); };
            windowWatcher.InCellEditWindowChanged += delegate { WatchInCellEdit(windowWatcher.InCellEditWindow); };
            windowWatcher.InCellEditFocused += FocusChanged;
            windowWatcher.FormulaBarFocused += FocusChanged;
            windowWatcher.MainWindowChanged += SubscribeBoundingRectangleProperty;
            _syncContext = syncContext;
        }

        private void SubscribeBoundingRectangleProperty(object sender, EventArgs args)
        {
            _syncContext.Post(delegate
            {
                if (_mainWindow != null)
                {
                    Automation.RemoveAutomationPropertyChangedEventHandler(_mainWindow, LocationChanged);
                }

                WindowWatcher windowWatcher = (WindowWatcher)sender;

                if (windowWatcher.MainWindow != IntPtr.Zero)
                {
                    // TODO: I see "Element not found" errors here...
                    //       There are different top-level windows in Excl 2013.
                    // _mainWindow = AutomationElement.FromHandle(windowWatcher.MainWindow);
                    // Automation.AddAutomationPropertyChangedEventHandler(_mainWindow, TreeScope.Element, LocationChanged, AutomationElement.BoundingRectangleProperty);
                }
            }, null);
        }

        void WatchFormulaBar(IntPtr hWnd)
        {
            _syncContext.Post(delegate
                {
                    Debug.Print("Will be watching Formula Bar: " + hWnd);
                    if (_formulaBar != null)
                    {
                        IntPtr hwndFormulaBar = (IntPtr)(int)_formulaBar.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);
                        if (hwndFormulaBar == hWnd)
                        {
                            // Window is fine - might be a text update
                            Debug.Print("Doing Update for FormulaBar");

                            UpdateEditState();
                            UpdateFormula();
                            return;
                        }
                        // TODO: Clean up
                        Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, _formulaBar, TextChanged);
                        _formulaBar = null;
                    }
                    _formulaBar = AutomationElement.FromHandle(hWnd);
                    UpdateEditState();

                    // CONSIDER: What when the formula resizes?
                    Automation.AddAutomationEventHandler(TextPattern.TextChangedEvent, _formulaBar, TreeScope.Element, TextChanged);
                }, null);
        }

        void WatchInCellEdit(IntPtr hWnd)
        {
            _syncContext.Post(delegate
                {
                    if (_inCellEdit != null)
                    {
                        // TODO: This is not a safe call if it has been destroyed
                        object objInCellHandle = _inCellEdit.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);
                        if (objInCellHandle != null && hWnd == (IntPtr)(int)objInCellHandle)
                        {
                            // Window is fine - might be a text update
                            Debug.Print("Doing Update for InCellEdit");
                            UpdateEditState();
                            UpdateFormula();
                            return;
                        }

                        // Change is upon us,
                        // out with the old...
                        Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, _inCellEdit, TextChanged);

                        _inCellEdit = null;
                    }

                    if (hWnd != IntPtr.Zero)
                    {
                        _inCellEdit = AutomationElement.FromHandle(hWnd);
                        UpdateEditState();
                        // CONSIDER: Can the incell box resize?
                        Automation.AddAutomationEventHandler(TextPattern.TextChangedEvent, _inCellEdit, TreeScope.Element, TextChanged);
                    }
                    UpdateEditState();
                }, null);
        }

        void TextChanged(object sender, AutomationEventArgs e)
        {
            Debug.WriteLine("! Active Text text changed. Is it the Formula Bar? {0}, Is it the In Cell Edit? {1}", sender.Equals(_formulaBar), sender.Equals(_inCellEdit));
            UpdateFormula();
        }

        void LocationChanged(object sender, AutomationPropertyChangedEventArgs e)
        {
            _syncContext.Post(delegate
                {
                    UpdateEditState(true);
                }, null);
        }

        void FocusChanged(object sender, EventArgs e)
        {
            UpdateEditState();
            UpdateFormula();
        }

        void UpdateEditState(bool moveOnly = false)
        {
            _syncContext.Post(delegate
                {
                    // TODO: This is not right yet - the list box might have the focus...
                    AutomationElement focused;
                    try
                    {
                        focused = AutomationElement.FocusedElement;
                    }
                    catch (ArgumentException aex)
                    {
                        Debug.Print("!!! ERROR: Failed to get Focused Element: " + aex.ToString());
                        // Not sure why I get this - sometimes with startup screen
                        return;
                    }
                    if (_formulaBar != null && _formulaBar.Equals(focused))
                    {
                        EditWindowBounds = (Rect)_formulaBar.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
                        IntPtr hwnd = (IntPtr)(int)_formulaBar.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);
                        var pt = Win32Helper.GetClientCursorPos(hwnd);
                        CaretPosition = new Point(pt.X, pt.Y);
                    }
                    else if (_inCellEdit != null && _inCellEdit.Equals(focused))
                    {
                        EditWindowBounds = (Rect)_inCellEdit.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
                        IntPtr hwnd = (IntPtr)(int)_inCellEdit.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);
                        var pt = Win32Helper.GetClientCursorPos(hwnd);
                        CaretPosition = new Point(pt.X, pt.Y);
                    }
                    else
                    {
                        // CurrentFormula = null;
                        CurrentPrefix = null;
                        Debug.Print("Don't have a focused text box to update.");
                    }

                    // As long as we have an InCellEdit, we are editing the formula...
                    IsEditingFormula = (_inCellEdit != null);

                    // TODO: Smarter notification...?
                    StateChanged(this, new StateChangeEventArgs(moveOnly ? StateChangeTypeEnum.Move : StateChangeTypeEnum.Undefined));
                }, null);
        }

        void UpdateFormula()
        {
            _syncContext.Post(delegate
                {
                    //            CurrentFormula = "";
                    CurrentPrefix = XlCall.GetFormulaEditPrefix();
                    StateChanged(this, new StateChangeEventArgs());
                }, null);
        }

        public void Dispose()
        {
            //_syncContext.Post(delegate 
            //    {
            Debug.Print("Disposing FormulaEditWatcher");
            if (_formulaBar != null)
            {
                Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, _formulaBar, TextChanged);
                _formulaBar = null;
            }
            if (_inCellEdit != null)
            {
                Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, _inCellEdit, TextChanged);
                _inCellEdit = null;
            }
            if (_mainWindow != null)
            {
                Automation.RemoveAutomationPropertyChangedEventHandler(_mainWindow, LocationChanged);
                _mainWindow = null;
            }
            //}, null);
        }
    }

    // The Popuplist is shown both for function selection, and for some argument selection.
    // We ignore this, and match purely on the text of the selected item.
    class PopupListWatcher : IDisposable
    {
        AutomationElement _popupList;
        AutomationElement _popupListList;
        AutomationElement _mainWindow;

        // NOTE: Event may be raised on a strange thread... (via Automation)
        public event EventHandler SelectedItemChanged = delegate { };

        public string SelectedItemText { get; set; }
        public Rect SelectedItemBounds { get; set; }

        private SynchronizationContext _syncContext;

        public PopupListWatcher(WindowWatcher windowWatcher, SynchronizationContext syncContext)
        {
            _syncContext = syncContext;
            windowWatcher.PopupListWindowChanged += delegate { WatchPopupList(windowWatcher.PopupListWindow); };
            windowWatcher.MainWindowChanged += SubscribeBoundingRectangleProperty;
        }

        void WatchPopupList(IntPtr hWnd)
        {
            _syncContext.Post(delegate
            {
                _popupList = AutomationElement.FromHandle(hWnd);
                Automation.AddStructureChangedEventHandler(_popupList, TreeScope.Element, PopupListStructureChangedHandler);
                // Automation.AddAutomationEventHandler(AutomationElement.v .AddStructureChangedEventHandler(_popupList, TreeScope.Element, PopupListStructureChangedHandler);
            } , null);
        }

        private void SubscribeBoundingRectangleProperty(object sender, EventArgs args)
        {
            _syncContext.Post(delegate
            {
                if (_mainWindow != null)
                {
                    Automation.RemoveAutomationPropertyChangedEventHandler(_mainWindow, LocationChanged);
                }

                WindowWatcher windowWatcher = (WindowWatcher)sender;

                if (windowWatcher.MainWindow != IntPtr.Zero)
                {
                    // TODO: I've seen an (ElementNotAvailableException) error here that 'the element is not available'.
                    _mainWindow = AutomationElement.FromHandle(windowWatcher.MainWindow);
                    Automation.AddAutomationPropertyChangedEventHandler(_mainWindow, TreeScope.Element, LocationChanged, AutomationElement.BoundingRectangleProperty);
                }
            }, null);
        }

        void LocationChanged(object sender, AutomationPropertyChangedEventArgs e)
        {
            _syncContext.Post(delegate
            {
                if (_popupListList != null)
                {
                    var selPat = _popupListList.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;
                    Automation.AddAutomationEventHandler(
                        SelectionItemPattern.ElementSelectedEvent, _popupListList, TreeScope.Children, PopupListElementSelectedHandler);

                    // Update the current selection, if any
                    var curSel = selPat.Current.GetSelection();
                    if (curSel.Length > 0)
                    {
                        try
                        {
                            UpdateSelectedItem(curSel[0]);
                        }
                        catch (Exception ex)
                        {
                            Debug.Print("Error during UpdateSelected! " + ex);
                        }
                    }
                }
            }, null);
        }

        // TODO: This should be exposed as an event and popup resize should be elsewhere
        private void PopupListStructureChangedHandler(object sender, StructureChangedEventArgs e)
        {
            _syncContext.Post(delegate
            {
                Debug.WriteLine(">>> PopupList structure changed - " + e.StructureChangeType.ToString());
                // CONSIDER: Others too?
                if (e.StructureChangeType == StructureChangeType.ChildAdded)
                {
                    var functionList = sender as AutomationElement;

                    var listRect = (Rect)functionList.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);

                    var listElement = functionList.FindFirst(TreeScope.Children, Condition.TrueCondition);
                    if (listElement != null)
                    {
                        _popupListList = listElement;

                        // TODO: Move this code!
                        TestMoveWindow(_popupListList, (int)listRect.X, (int)listRect.Y);
                        TestMoveWindow(functionList, 0, 0);

                        var selPat = listElement.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;
                        Automation.AddAutomationEventHandler(
                            SelectionItemPattern.ElementSelectedEvent, listElement, TreeScope.Children, PopupListElementSelectedHandler);

                        // Update the current selection, if any
                        var curSel = selPat.Current.GetSelection();
                        if (curSel.Length > 0)
                        {
                            try
                            {
                                UpdateSelectedItem(curSel[0]);
                            }
                            catch (Exception ex)
                            {
                                Debug.Print("Error during UpdateSelected! " + ex);
                            }
                        }
                    }
                    else
                    {
                        Debug.WriteLine("ERROR!!! Structure changed but no children anymore.");
                    }
                }
                else if (e.StructureChangeType == StructureChangeType.ChildRemoved)
                {
                    if (_popupListList != null)
                    {
                        Automation.RemoveAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent, _popupListList, PopupListElementSelectedHandler);
                        _popupListList = null;
                    }
                    SelectedItemText = String.Empty;
                    SelectedItemBounds = Rect.Empty;
                    SelectedItemChanged(this, EventArgs.Empty);
                }
            }, null);
        }

        private void TestMoveWindow(AutomationElement listWindow, int xOffset, int yOffset)
        {
            var hwndList = (IntPtr)(int)(listWindow.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty));
            var listRect = (Rect)listWindow.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
            Debug.Print("Moving from {0}, {1}", listRect.X, listRect.Y);
            Win32Helper.MoveWindow(hwndList, (int)listRect.X - xOffset, (int)listRect.Y - yOffset, (int)listRect.Width, 2 * (int)listRect.Height, true);
        }

        // Must run on the UI thread.
        private void UpdateSelectedItem(AutomationElement automationElement)
        {
            SelectedItemText = (string)automationElement.GetCurrentPropertyValue(AutomationElement.NameProperty);
            SelectedItemBounds = (Rect)automationElement.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
            SelectedItemChanged(this, EventArgs.Empty);
        }

        void PopupListElementSelectedHandler(object sender, AutomationEventArgs e)
        {
            _syncContext.Post(delegate
            {
                Debug.Print("### Thread receiving PopupListElementSelectedHandler: " + Thread.CurrentThread.ManagedThreadId);
                UpdateSelectedItem(sender as AutomationElement);
            }, null);
        }

        public void Dispose()
        {
            //_syncContext.Post(delegate
            //{
            Debug.Print("Disposing PopupListWatcher");
            if (_popupList != null)
            {
                Automation.RemoveStructureChangedEventHandler(_popupList, PopupListStructureChangedHandler);
                _popupList = null;

                if (_popupListList != null)
                {
                    Automation.RemoveAutomationEventHandler(SelectionItemPattern.ElementSelectedEvent, _popupListList, PopupListElementSelectedHandler);
                    _popupListList = null;
                }
            }
            //}, null);
        }
    }

    class SelectDataSourceWatcher : IDisposable
    {
        private SynchronizationContext _syncContext;

        public SelectDataSourceWatcher(WindowWatcher windowWatcher, SynchronizationContext syncContext)
        {
            _syncContext = syncContext;
            
        }

        public void Dispose()
        {

        }
    }
}
