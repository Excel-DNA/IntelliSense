using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using Point = System.Drawing.Point;

namespace ExcelDna.IntelliSense
{
    class WindowWatcher : IDisposable
    {
        WinEventHook _windowStateChangeHook;

        public IntPtr MainWindow { get; private set; }
        public IntPtr FormulaBarWindow { get; private set; }
        public IntPtr InCellEditWindow { get; private set; }
        public IntPtr PopupListWindow { get; private set; }
        public IntPtr PopupListList { get; private set; }

        public event EventHandler FormulaBarWindowChanged = delegate { };
        public event EventHandler FormulaBarFocused = delegate { };
        public event EventHandler InCellEditWindowChanged = delegate { };
        public event EventHandler InCellEditFocused = delegate { };
        public event EventHandler MainWindowChanged = delegate { };
        public event EventHandler PopupListWindowChanged = delegate { };   // Might start off with nothing. Changes at most once.
        public event EventHandler PopupListListChanged = delegate { };   // Might start off with nothing. Changes at most once.

        public WindowWatcher(string xllName)
        {
            Debug.Print("### WindowWatcher created on thread: " + Thread.CurrentThread.ManagedThreadId);

            // Using WinEvents instead of Automation so that we can watch top-level changes, but only from the right process.

            _windowStateChangeHook = new WinEventHook(WindowStateChange,
                WinEventHook.WinEvent.EVENT_OBJECT_CREATE, WinEventHook.WinEvent.EVENT_OBJECT_FOCUS, xllName);
        }

        public void TryInitialize()
        {

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
            Debug.Print("### Thread receiving WindowStateChange: " + Thread.CurrentThread.ManagedThreadId);
            switch (Win32Helper.GetClassName(hWnd))
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

        public void Dispose()
        {
            if (_windowStateChangeHook != null)
            {
                _windowStateChangeHook.Stop();
                _windowStateChangeHook = null;
            }
        }
    }

    // We want to know whether to show the function entry help
    // For now we ignore whether another ToolTip is being shown, and just use the formula edit state.
    class FormulaEditWatcher : IDisposable
    {
        AutomationElement _formulaBar;
        AutomationElement _inCellEdit;

        public event EventHandler StateChanged = delegate { };

        public bool IsEditingFormula { get; set; }
//        public string CurrentFormula { get; set; } // Easy to get, but we don't need it
        public string CurrentPrefix { get; set; }    // Null if not editing
        // We don't really care whether it is the formula bar or in-cell, 
        // we just need to get the right window handle 
        public Rect EditWindowBounds { get; set; }
        public Point CaretPosition { get; set; }

        private DispatcherThread _dispatcherThread;

        public FormulaEditWatcher(WindowWatcher windowWatcher, DispatcherThread dispatcher)
        {
            windowWatcher.FormulaBarWindowChanged += delegate { WatchFormulaBar(windowWatcher.FormulaBarWindow); };
            windowWatcher.InCellEditWindowChanged += delegate { WatchInCellEdit(windowWatcher.InCellEditWindow); };
            windowWatcher.InCellEditFocused += FocusChanged;
            windowWatcher.FormulaBarFocused += FocusChanged;
            _dispatcherThread = dispatcher;
        }

        void WatchFormulaBar(IntPtr hWnd)
        {
            _dispatcherThread.Invoke(() =>
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
                });
        }
             
        void WatchInCellEdit(IntPtr hWnd)
        {
            _dispatcherThread.Invoke(() =>
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
                });
        }

        void TextChanged(object sender, AutomationEventArgs e)
        {
            Debug.WriteLine("! Active Text text changed. Is it the Formula Bar? {0}, Is it the In Cell Edit? {1}", sender.Equals(_formulaBar), sender.Equals(_inCellEdit));
            UpdateFormula();
        }

        void FocusChanged(object sender, EventArgs e)
        {
            UpdateEditState();
            UpdateFormula();
        }

        void UpdateEditState()
        {
            _dispatcherThread.Invoke(() =>
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
                    StateChanged(this, EventArgs.Empty);
                });
        }

        void UpdateFormula()
        {
            _dispatcherThread.Invoke(() =>
                {
                    //            CurrentFormula = "";
                    CurrentPrefix = XlCall.GetFormulaEditPrefix();
                    StateChanged(this, EventArgs.Empty);
                });
        }

        public void Dispose()
        {
            _dispatcherThread.Invoke(() =>
                {
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
                });
        }
    }

    // The Popuplist is shown both for function selection, and for some argument selection.
    // We ignore this, and match purely on the text of the selected item.
    class PopupListWatcher : IDisposable
    {
        AutomationElement _popupList;
        AutomationElement _popupListList;

        // NOTE: Event may be raised on a strange thread... (via Automation)
        public event EventHandler SelectedItemChanged = delegate { };

        public string SelectedItemText { get; set; }
        public Rect SelectedItemBounds { get; set; }

        public PopupListWatcher(WindowWatcher windowWatcher)
        {
            windowWatcher.PopupListWindowChanged += delegate { WatchPopupList(windowWatcher.PopupListWindow); };
        }

        void WatchPopupList(IntPtr hWnd)
        {
            _popupList = AutomationElement.FromHandle(hWnd);
            Automation.AddStructureChangedEventHandler(_popupList, TreeScope.Element, PopupListStructureChangedHandler);
            // Automation.AddAutomationEventHandler(AutomationElement.v .AddStructureChangedEventHandler(_popupList, TreeScope.Element, PopupListStructureChangedHandler);
        }

        // TODO: This should be exposed as an event and popup resize should be elsewhere
        private void PopupListStructureChangedHandler(object sender, StructureChangedEventArgs e)
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
                    TestMoveWindow(functionList, 0,0);

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
                    Automation.RemoveStructureChangedEventHandler(_popupListList, PopupListElementSelectedHandler);
                    _popupListList = null;
                }
                SelectedItemText = String.Empty;
                SelectedItemBounds = Rect.Empty;
                SelectedItemChanged(this, EventArgs.Empty);
            }
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
            Debug.Print("### Thread receiving PopupListElementSelectedHandler: " + Thread.CurrentThread.ManagedThreadId);
            UpdateSelectedItem(sender as AutomationElement);
        }

        public void Dispose()
        {
            Debug.Print("Disposing PopupListWatcher");
            if (_popupList != null)
            {
                Automation.RemoveStructureChangedEventHandler(_popupList, PopupListStructureChangedHandler);
                _popupList = null;

                if (_popupListList != null)
                {
                    Automation.RemoveStructureChangedEventHandler(_popupListList, PopupListElementSelectedHandler);
                    _popupListList = null;
                }
            }
        }
    }
}
