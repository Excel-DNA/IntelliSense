using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Automation.Text;
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

        public event EventHandler FormulaBarWindowChanged = delegate { };
        public event EventHandler FormulaBarFocused = delegate { };
        public event EventHandler InCellEditWindowChanged = delegate { };
        public event EventHandler InCellEditFocused = delegate { };
        public event EventHandler MainWindowChanged = delegate { };
        public event EventHandler PopupListWindowChanged = delegate { };   // Might start off with nothing. Changes at most once.

        public WindowWatcher()
        {
            Debug.Print("### WindowWatcher created on thread: " + Thread.CurrentThread.ManagedThreadId);

            // Using WinEvents instead of Automation so that we can watch top-level changes, but only from the right process.

            _windowStateChangeHook = new WinEventHook(WindowStateChange,
                WinEventHook.WinEvent.EVENT_OBJECT_CREATE, WinEventHook.WinEvent.EVENT_OBJECT_FOCUS);
        }

        public void Initialize()
        {
            AutomationElement focused = AutomationElement.FocusedElement;
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
            // Debug.Print("### Thread receiving WindowStateChange: " + Thread.CurrentThread.ManagedThreadId);
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
        TextPattern _activeTextPattern;

        public event EventHandler StateChanged = delegate { };

        public bool IsEditingFormula { get; set; }
        public string CurrentFormula { get; set; }
        public string CurrentPrefix { get; set; }
        // We don't really care whether it is the formula bar or in-cell, 
        // we just need to get the right window handle 
        public Rect EditWindowBounds { get; set; }
        public System.Drawing.Point CaretPosition { get; set; }

        public FormulaEditWatcher(WindowWatcher windowWatcher)
        {
            windowWatcher.FormulaBarWindowChanged += delegate { WatchFormulaBar(windowWatcher.FormulaBarWindow); };
            windowWatcher.InCellEditWindowChanged += delegate { WatchInCellEdit(windowWatcher.InCellEditWindow); };
            windowWatcher.InCellEditFocused += FocusChanged;
            windowWatcher.FormulaBarFocused += FocusChanged;
        }

        void WatchFormulaBar(IntPtr hWnd)
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
        }

        void WatchInCellEdit(IntPtr hWnd)
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
        }

        void TextChanged(object sender, AutomationEventArgs e)
        {
            Debug.WriteLine("! Active Text text changed. Is it the Formula Bar? {0}, Is it the In Cell Edit? {1}", sender.Equals(_formulaBar), sender.Equals(_inCellEdit));
            UpdateFormula();
        }

        void FocusChanged(object sender, EventArgs e)
        {
            _activeTextPattern = null;
            UpdateEditState();
            UpdateFormula();
        }

        void UpdateEditState()
        {
            // TODO: This is not right yet - the list box might have the focus...

            if (_formulaBar != null && _formulaBar.Equals(AutomationElement.FocusedElement))
            {
                EditWindowBounds = (Rect)_formulaBar.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
                IntPtr hwnd = (IntPtr)(int)_formulaBar.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);
                var pt = Win32Helper.GetClientCursorPos(hwnd);
                CaretPosition = new System.Drawing.Point(pt.X, pt.Y);
            }
            else if (_inCellEdit != null && _inCellEdit.Equals(AutomationElement.FocusedElement))
            {
                EditWindowBounds = (Rect)_inCellEdit.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
                IntPtr hwnd = (IntPtr)(int)_inCellEdit.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);
                var pt = Win32Helper.GetClientCursorPos(hwnd);
                CaretPosition = new Point(pt.X, pt.Y);
            }
            else
            {
                CurrentFormula = null;
                CurrentPrefix = null;
                Debug.Print("Don't have a focused text box to update.");
            }

            // As long as we have an InCellEdit, we are editing the formula...
            IsEditingFormula = (_inCellEdit != null);

            // TODO: Smarter notification...?
            StateChanged(this, EventArgs.Empty);
        }

        void UpdateActivePattern()
        {
            if (_formulaBar != null && _formulaBar.Equals(AutomationElement.FocusedElement))
            {
                Debug.Print("FormulaBar has focus - ClassName: " + _formulaBar.GetCurrentPropertyValue(AutomationElement.ClassNameProperty));
                var supportsText = (bool)_formulaBar.GetCurrentPropertyValue(AutomationElement.IsTextPatternAvailableProperty);
                if (!supportsText)
                {
                    Debug.Print("No support for TextPattern in FormulaBar!?");
                    return;
                }
                _activeTextPattern = _formulaBar.GetCurrentPattern(TextPattern.Pattern) as TextPattern;
            }
            else if (_inCellEdit != null /* && _inCellEdit.Equals(AutomationElement.FocusedElement)*/)
            {

                Debug.Print("InCell has focus - ClassName: " + _inCellEdit.GetCurrentPropertyValue(AutomationElement.ClassNameProperty));
                Debug.Print("Checking InCellEdit text pattern availability property on thread {0}", Thread.CurrentThread.ManagedThreadId);

                var supportsText = (bool)_inCellEdit.GetCurrentPropertyValue(AutomationElement.IsTextPatternAvailableProperty);
                if (!supportsText)
                {
                    // Attempt to refresh InCell Automation Element. It is half-baked.
                    IntPtr hwndInCell = (IntPtr)(int)_inCellEdit.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);
                    Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, _inCellEdit, TextChanged);
                    _inCellEdit = AutomationElement.FromHandle(hwndInCell);
                    supportsText = (bool)_inCellEdit.GetCurrentPropertyValue(AutomationElement.IsTextPatternAvailableProperty);
                }

                if (!supportsText)
                {
                    Debug.Print("Could not fix InCellEdit yet... Abandoning update");
                    return;
                }
                _activeTextPattern = _inCellEdit.GetCurrentPattern(TextPattern.Pattern) as TextPattern;
            }
            else
            {
                Debug.Print("Don't have a focused text box to update.");
                return;
            }

        }

        void UpdateFormula()
        {
            UpdateActivePattern();
            if (_activeTextPattern == null) return;
            try
            {
                TextPatternRange fullRange = _activeTextPattern.DocumentRange;
                CurrentFormula = fullRange.GetText(-1); // Get all text

                TextPatternRange[] selections = _activeTextPattern.GetSelection();
                if (selections.Length == 1)
                {
                    TextPatternRange curSel = selections[0];
                    TextPatternRange selClone = curSel.Clone();
                    selClone.MoveEndpointByRange(TextPatternRangeEndpoint.Start, fullRange, TextPatternRangeEndpoint.Start);
                    CurrentPrefix = selClone.GetText(-1);

                    Debug.Print("Formula: {0}, Prefix: {1}", CurrentFormula, CurrentPrefix);
                    StateChanged(this, EventArgs.Empty);
                    // Other bits in Win7 api:
                    //    // UIA_CaretPositionAttributeId
                    //    // UIA_SelectionActiveEndAttributeId
                    //    // UIA_IsActiveAttributeId
                }
            }
            catch (ElementNotEnabledException)
            {
                Debug.Print("Element not enabled!?");
                StateChanged(this, EventArgs.Empty);
            }
        }

        public void Dispose()
        {
            Debug.Print("Disposing FormulaEditWatcher");
            _activeTextPattern = null;
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
            // CONSIDER: Others too?
            if (e.StructureChangeType == StructureChangeType.ChildAdded)
            {
                var functionList = sender as AutomationElement;
                var listRect = (Rect)functionList.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);

                var listElement = functionList.FindFirst(TreeScope.Children, Condition.TrueCondition);
                if (listElement != null)
                {
                    _popupListList = listElement;
                    // TestMoveWindow(_popupListList, (int)listRect.X, (int)listRect.Y);

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
