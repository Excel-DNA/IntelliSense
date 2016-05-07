using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Automation;

namespace ExcelDna.IntelliSense
{
    // We want to know whether to show the function entry help
    // For now we ignore whether another ToolTip is being shown, and just use the formula edit state.
    // CONSIDER: Should we watch the in-cell edit box and the formula edit control separately?
    class FormulaEditWatcher : IDisposable
    {
        enum FormulaEditFocus
        {
            None = 0,
            FormulaBar = 1,
            InCellEdit = 2
        }

        public enum StateChangeType
        {
            Multiple,
            Move,
            TextChange
        }

        public class StateChangeEventArgs : EventArgs
        {
            public static new StateChangeEventArgs Empty = new StateChangeEventArgs();

            public StateChangeType StateChangeType { get; private set; }

            public StateChangeEventArgs(StateChangeType? stateChangeType = null)
            {
                StateChangeType = stateChangeType ?? StateChangeType.Multiple;
            }
        }

        // NOTE: Our event will always be raised on the _syncContextAuto thread (CONSIDER: Does this help?)
        public event EventHandler<StateChangeEventArgs> StateChanged;

        public bool IsEditingFormula { get; set; }
        public string CurrentPrefix { get; set; }    // Null if not editing
        // We don't really care whether it is the formula bar or in-cell, 
        // we just need to get the right window handle 
        public Rect EditWindowBounds { get; set; }
        public Point CaretPosition { get; set; }
        public IntPtr FormulaEditWindow
        {
            get
            {
                switch (_formulaEditFocus)
                {
                    case FormulaEditFocus.None:
                        return IntPtr.Zero;
                    case FormulaEditFocus.FormulaBar:
                        return _hwndFormulaBar;
                    case FormulaEditFocus.InCellEdit:
                        return _hwndInCellEdit;
                    default:
                        throw new InvalidOperationException("Invalid FormulaEditWatcher.FormulaEditFocus");
                }
            }
        }

        SynchronizationContext _syncContextAuto;
        WindowWatcher _windowWatcher;

        IntPtr            _hwndFormulaBar;
        IntPtr            _hwndInCellEdit;
        AutomationElement _formulaBar;
        AutomationElement _inCellEdit;
        FormulaEditFocus  _formulaEditFocus;

        bool _enableFormulaPolling = false;
        Timer _formulaPollingTimer;

        public FormulaEditWatcher(WindowWatcher windowWatcher, SynchronizationContext syncContextAuto)
        {
            _syncContextAuto = syncContextAuto;
            _windowWatcher = windowWatcher;
            _windowWatcher.FormulaBarWindowChanged += _windowWatcher_FormulaBarWindowChanged;
            _windowWatcher.InCellEditWindowChanged += _windowWatcher_InCellEditWindowChanged;

            // Focus event handler works beautifully, but Breaks the PopupList SelectedItemChange event handler ... !?
            //_syncContextAuto.Post(_ =>
            //{
            //    try
            //    {
            //        Automation.AddAutomationFocusChangedEventHandler(FocusChangedEventHandler);
            //        Logger.WindowWatcher.Verbose("FormulaEditWatcher Focus event handler added");
            //    }
            //    catch (Exception ex)
            //    {
            //        Logger.WindowWatcher.Warn($"FormulaEditWatcher Error adding focus event handler {ex}");
            //    }
            //}, null);
        }

        // Runs on the Automation thread
        void _windowWatcher_FormulaBarWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs e)
        {
            switch (e.Type)
            {
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Create:
                    // CONSIDER: Is this too soon to set up the AutomationElement ??
                    SetEditWindow(e.WindowHandle, ref _hwndFormulaBar, ref _formulaBar);
                    UpdateEditState();
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Destroy:
                    //if (_formulaEditFocus == FormulaEditFocus.FormulaBar)
                    //{
                    //    _formulaEditFocus = FormulaEditFocus.None;
                    //    UpdateEditState();
                    //}
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Focus:
                    if (_formulaEditFocus != FormulaEditFocus.FormulaBar)
                    {
                        Logger.WindowWatcher.Verbose($"FormulaEdit - FormulaBar focus");
                        _formulaEditFocus = FormulaEditFocus.FormulaBar;
                        UpdateFormulaPolling();
                        UpdateEditState();
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Show:
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Hide:
                    if (_formulaEditFocus == FormulaEditFocus.FormulaBar)
                    {
                        Logger.WindowWatcher.Verbose($"FormulaEdit - FormulaBar hidden");
                        _formulaEditFocus = FormulaEditFocus.None;
                        UpdateFormulaPolling();
                        UpdateEditState();
                    }
                    break;
                default:
                    throw new ArgumentOutOfRangeException("Unexpected Window Change Type", "e.Type");
            }
        }

        // Runs on the Automation thread
        void _windowWatcher_InCellEditWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs e)
        {
            switch (e.Type)
            {
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Create:
                    // CONSIDER: Is this too soon to set up the AutomationElement ??
                    // TODO: Yes - need to do AutomationElement later (new Window does not have TextPattern ready)
                    SetEditWindow(e.WindowHandle, ref _hwndInCellEdit, ref _inCellEdit);
                    UpdateEditState();
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Destroy:
                    //if (_formulaEditFocus == FormulaEditFocus.InCellEdit)
                    //{
                    //    _formulaEditFocus = FormulaEditFocus.None;
                    //    UpdateEditState();
                    //}
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Focus:
                    if (_formulaEditFocus != FormulaEditFocus.InCellEdit)
                    {
                        Logger.WindowWatcher.Verbose($"FormulaEdit - InCellEdit focus");
                        _formulaEditFocus = FormulaEditFocus.InCellEdit;
                        UpdateFormulaPolling();
                        UpdateEditState();
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Show:
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Hide:
                    // NOTE: Very confusing - under Excel 2010 InCellEdit gets hidden immediately, and never shown again....
                    //if (_formulaEditFocus == FormulaEditFocus.InCellEdit)
                    //{
                    //    Logger.WindowWatcher.Verbose($"FormulaEdit - InCellEdit hidden");
                    //    _formulaEditFocus = FormulaEditFocus.None;
                    //    UpdateFormulaPolling();
                    //    UpdateEditState();
                    //}
                    break;
                default:
                    throw new ArgumentOutOfRangeException("Unexpected Window Change Type", "e.Type");
            }
        }

        void SetEditWindow(IntPtr newWindowHandle, ref IntPtr hwnd, ref AutomationElement element)
        {
            if (hwnd != newWindowHandle)
            {
                if (hwnd == IntPtr.Zero)
                {
                    Logger.WindowWatcher.Info($"FormulaEdit SetEditWindow - Initialize {newWindowHandle}");    // Could be change of Excel window .... ?
                }
                else
                {
                    Logger.WindowWatcher.Info($"FormulaEdit SetEditWindow - Change from {hwnd} to {newWindowHandle}");
                }

                if (element != null)
                {
                    try
                    {
                        UninstallTextChangeMonitor(element);
                        Logger.WindowWatcher.Verbose($"FormulaEdit Uninstalled event handlers for {hwnd}");
                    }
                    catch (Exception ex)
                    {
                        Logger.WindowWatcher.Warn($"FormulaEdit Error uninstalling event handlers for {hwnd}: {ex}");
                    }
                }
            }
            else
            {
                // same window - ignore
                return;
            }

            // Setting the out parameters
            hwnd = newWindowHandle;
            element = AutomationElement.FromHandle(newWindowHandle);

            try
            {
                InstallTextChangeMonitor(element);
                Logger.WindowWatcher.Verbose($"FormulaEdit Installed event handlers for {newWindowHandle}");
            }
            catch (Exception ex)
            {
                Logger.WindowWatcher.Warn($"FormulaEdit Error installing event handlers for {newWindowHandle}: {ex}");
            }
        }

        void InstallTextChangeMonitor(AutomationElement element)
        {
            // Either add a UIAutomation event handler, or start our polling
            // CONSIDER: Try Keyboard hook or other Windows message watcher of some sort?
            bool isTextPatternAvailable = (bool)element.GetCurrentPropertyValue(AutomationElement.IsTextPatternAvailableProperty);

            if (isTextPatternAvailable)
            {
                Logger.WindowWatcher.Info("FormulaEdit TextPattern adding change handler");
                Automation.AddAutomationEventHandler(TextPattern.TextChangedEvent, element, TreeScope.Element, TextChanged);
                _enableFormulaPolling = false;
            }
            else
            {
                Logger.WindowWatcher.Info("FormulaEdit TextPattern not available - enabling polling");
                _enableFormulaPolling = true;   // Has an effect when focus changes
                UpdateFormulaPolling();
            }
        }

        void UninstallTextChangeMonitor(AutomationElement element)
        {
            bool isTextPatternAvailable = (bool)element.GetCurrentPropertyValue(AutomationElement.IsTextPatternAvailableProperty);
            if (isTextPatternAvailable)
            {
                Logger.WindowWatcher.Info("FormulaEdit TextPattern removing change handler");
                Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, element, TextChanged);
            }
        }

        // Breaks the PopupList SelectedItemChange event handler ... !?
        //// Expected on some automation event handler thread
        //void FocusChangedEventHandler(object sender, AutomationFocusChangedEventArgs e)
        //{
        //    return;
        //    Logger.WindowWatcher.Verbose($"++++> Focus Changed! ({_formulaEditFocus})");
        //    if (_formulaEditFocus == FormulaEditFocus.None)
        //        return;

        //    var formulaPrefix = XlCall.GetFormulaEditPrefix();
        //    if (formulaPrefix == null)
        //    {
        //        // We need to update
        //        _formulaEditFocus = FormulaEditFocus.None;
        //        UpdateFormulaPolling();
        //        UpdateEditState();
        //    }
        //}

        // Threading... ???
        void UpdateFormulaPolling()
        {
            if (_enableFormulaPolling)
            {
                if (_formulaEditFocus != FormulaEditFocus.None)
                {
                    if (_formulaPollingTimer == null)
                        _formulaPollingTimer = new Timer(FormulaPollingCallback);
                    _formulaPollingTimer.Change(100, 100);
                }
                else // no focus
                {
                    if (_formulaPollingTimer != null)
                        _formulaPollingTimer.Change(-1, -1);
                }
            }
            else if (_formulaPollingTimer != null)
            {
                _formulaPollingTimer.Dispose();
                _formulaPollingTimer = null;
            }
        }

        // Called on some ThreadPool thread...
        void FormulaPollingCallback(object _unused_)
        {
            // Logger.WindowWatcher.Verbose($"FormulaEdit - FormulaPollingCallback");

            // Check for Disposed already
            // TODO: Check again whether System.Timers.Timer can fire after Dispose()...
            if (_formulaPollingTimer == null)
                return;

            var oldPrefix = CurrentPrefix;
            var newPrefix = XlCall.GetFormulaEditPrefix();
            if (oldPrefix != newPrefix)
            {
                // CONSIDER: What to do if newPrefix is null (not editing...)
                CurrentPrefix = newPrefix;
                OnStateChanged(new StateChangeEventArgs(StateChangeType.TextChange));
            }
        }

        // Runs on an Automation event thread
        // CONSIDER: With WinEvents we could get notified from main thread ... ?
        void TextChanged(object sender, AutomationEventArgs e)
        {
            // Debug.Print($">>>> FormulaEditWatcher.TextChanged on thread {Thread.CurrentThread.ManagedThreadId}");
            Logger.WindowWatcher.Verbose($"FormulaEdit - Text changed in {(sender.Equals(_formulaBar) ? "FormulaBar" : (sender.Equals(_inCellEdit) ? "InCellEdit" : "UNKNOWN"))}");
            UpdateFormula(textChangedOnly: true);
        }

        // Runs on an Automation event thread
        // CONSIDER: With WinEvents we could get notified from main thread ... ?
        void LocationChanged(object sender, AutomationPropertyChangedEventArgs e)
        {
            // Debug.Print($">>>> FormulaEditWatcher.LocationChanged on thread {Thread.CurrentThread.ManagedThreadId}");
            Logger.WindowWatcher.Verbose($"FormulaEdit - Location changed");
            UpdateEditState(true);
        }

        void UpdateEditState(bool moveOnly = false)
        {
            Logger.WindowWatcher.Verbose("> FormulaEdit UpdateEditState - Posted");
            _syncContextAuto.Post(moveOnlyObj =>
                {
                    Logger.WindowWatcher.Verbose($"FormulaEdit UpdateEditState - Focus: {_formulaEditFocus}");
                    //// TODO: This is not right yet - the list box might have the focus... ?
                    //AutomationElement focused;
                    //try
                    //{
                    //    focused = AutomationElement.FocusedElement;
                    //}
                    //catch (ArgumentException aex)
                    //{
                    //    Debug.Print($"!!! ERROR: Failed to get Focused Element: {aex}");
                    //    // Not sure why I get this - sometimes with startup screen
                    //    return;
                    //}
                    AutomationElement focusedEdit = null;
                    bool prefixChanged = false;
                    if (_formulaEditFocus == FormulaEditFocus.FormulaBar)
                    {
                        focusedEdit = _formulaBar;
                    }
                    else if (_formulaEditFocus == FormulaEditFocus.InCellEdit)
                    {
                        focusedEdit = _inCellEdit;
                    }
                    else
                    {
                        // Neither have the focus, so we don't update anything
                        Logger.WindowWatcher.Verbose("FormulaEdit UpdateEditState End formula editing");
                        CurrentPrefix = null;
                        IsEditingFormula = false;
                        prefixChanged = true;
                        // Debug.Print("Don't have a focused text box to update.");
                    }

                    if (focusedEdit != null)
                    {
                        EditWindowBounds = (Rect)focusedEdit.GetCurrentPropertyValue(AutomationElement.BoundingRectangleProperty);
                        IntPtr hwnd = (IntPtr)(int)focusedEdit.GetCurrentPropertyValue(AutomationElement.NativeWindowHandleProperty);

                        var pt = Win32Helper.GetClientCursorPos(hwnd);
                        CaretPosition = new Point(pt.X, pt.Y);
                        IsEditingFormula = true;
                        var newPrefix = XlCall.GetFormulaEditPrefix();  // What thread do we have to use here ...?
                        if (CurrentPrefix != newPrefix)
                        {
                            CurrentPrefix = newPrefix;
                            prefixChanged = true;
                        }
                        Logger.WindowWatcher.Verbose($"FormulaEdit UpdateEditState Formula editing: CurrentPrefix {CurrentPrefix}, EditWindowBounds: {EditWindowBounds}, CaretPosition {CaretPosition}");
                    }

                    // TODO: Smarter notification...?
                    OnStateChanged(new StateChangeEventArgs(((bool)moveOnlyObj && !prefixChanged) ? StateChangeType.Move : StateChangeType.Multiple));
                }, moveOnly);
        }

        void UpdateFormula(bool textChangedOnly = false)
        {
            Logger.WindowWatcher.Verbose($">>>> FormulaEditWatcher.UpdateFormula on thread {Thread.CurrentThread.ManagedThreadId}");
            CurrentPrefix = XlCall.GetFormulaEditPrefix();  // What thread do we have to use here ...?
            Logger.WindowWatcher.Verbose($">>>> FormulaEditWatcher.UpdateFormula CurrentPrefix: {CurrentPrefix}");
            OnStateChanged(textChangedOnly ? new StateChangeEventArgs(StateChangeType.TextChange) : StateChangeEventArgs.Empty);
        }

        // We ensure that our event is raised on the Automation thread .. (Eases concurrency issues in handling it, though it will get passed on to the main thread...)
        void OnStateChanged(StateChangeEventArgs stateChangeEventArgs)
        {
            _syncContextAuto.Post(args => StateChanged?.Invoke(this, (StateChangeEventArgs)args), stateChangeEventArgs);
        }

        public void Dispose()
        {
            if (_formulaPollingTimer != null)
            {
                _formulaPollingTimer.Dispose();
            }
            // Not sure we need this:
            _windowWatcher.FormulaBarWindowChanged -= _windowWatcher_FormulaBarWindowChanged;
            _windowWatcher.InCellEditWindowChanged -= _windowWatcher_InCellEditWindowChanged;
//            _windowWatcher.MainWindowChanged -= _windowWatcher_MainWindowChanged;

            _syncContextAuto.Post(delegate 
            {
                Debug.Print("Disposing FormulaEditWatcher");
                //Automation.RemoveAutomationFocusChangedEventHandler(FocusChangedEventHandler);
                //if (_formulaBar != null)
                //{
                //    Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, _formulaBar, TextChanged);
                //    _formulaBar = null;
                //}
                if (_inCellEdit != null)
                {
                    Automation.RemoveAutomationEventHandler(TextPattern.TextChangedEvent, _inCellEdit, TextChanged);
                    _inCellEdit = null;
                }
                //if (_mainWindow != null)
                //{
                //    Automation.RemoveAutomationPropertyChangedEventHandler(_mainWindow, LocationChanged);
                //    _mainWindow = null;
                //}
            }, null);
        }
    }
}
