using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;

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

        readonly static StateChangeEventArgs s_stateChangeMultiple = new StateChangeEventArgs(StateChangeType.Multiple);
        readonly static StateChangeEventArgs s_stateChangeMove = new StateChangeEventArgs(StateChangeType.Move);
        readonly static StateChangeEventArgs s_stateChangeTextChange = new StateChangeEventArgs(StateChangeType.TextChange);

        // NOTE: Our event will always be raised on the _syncContextAuto thread (CONSIDER: Does this help?)
        public event EventHandler<StateChangeEventArgs> StateChanged;

        public bool IsEditingFormula { get; set; }
        public string CurrentPrefix { get; set; }    // Null if not editing
        // We don't really care whether it is the formula bar or in-cell, 
        // we just need to get the right window handle 
        public Rect EditWindowBounds { get; set; }

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

        readonly SynchronizationContext _syncContextAuto;
        readonly SynchronizationContext _syncContextMain;

        readonly WindowWatcher _windowWatcher;          // Passed in
        WindowLocationWatcher  _windowLocationWatcher;  // Managed here

        IntPtr            _hwndFormulaBar;
        IntPtr            _hwndInCellEdit;
        FormulaEditFocus  _formulaEditFocus;

        public FormulaEditWatcher(WindowWatcher windowWatcher, SynchronizationContext syncContextAuto, SynchronizationContext syncContextMain)
        {
            _syncContextAuto = syncContextAuto;
            _syncContextMain = syncContextMain;
            _windowWatcher = windowWatcher;
            _windowWatcher.FormulaBarWindowChanged += _windowWatcher_FormulaBarWindowChanged;
            _windowWatcher.InCellEditWindowChanged += _windowWatcher_InCellEditWindowChanged;
        }

        // Runs on the Automation thread
        void _windowWatcher_FormulaBarWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs e)
        {
            switch (e.Type)
            {
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Create:
                    if (e.ObjectId == WindowWatcher.WindowChangedEventArgs.ChangeObjectId.Self)
                    {
                        SetEditWindow(e.WindowHandle, ref _hwndFormulaBar);
                        UpdateEditState();
                    }
                    else if (e.ObjectId == WindowWatcher.WindowChangedEventArgs.ChangeObjectId.Caret)
                    {
                        // We expect this on every text change
                        // NOTE: Not anymore after some Excel / Windows update
                        UpdateEditStateDelayed();
                    }
                    else
                    {
                        Debug.Print($"### Unexpected WindowsChanged object id: {e.ObjectId}");
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Destroy:
                    // We expect this for every text change, but ignore since we react to the Create event
                    //if (_formulaEditFocus == FormulaEditFocus.FormulaBar)
                    //{
                    //    _formulaEditFocus = FormulaEditFocus.None;
                    //    UpdateEditState();
                    //}
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Focus:
                    if (_formulaEditFocus != FormulaEditFocus.FormulaBar)
                    {
                        Logger.WindowWatcher.Verbose($"FormulaEdit - FormulaBar Focus");
                        if (e.WindowHandle != _hwndFormulaBar)
                        {
                            // We never saw the Create...
                            SetEditWindow(e.WindowHandle, ref _hwndFormulaBar);
                        }
                        _formulaEditFocus = FormulaEditFocus.FormulaBar;
                        UpdateEditState();
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Unfocus:
                    if (_formulaEditFocus == FormulaEditFocus.FormulaBar)
                    {
                        Logger.WindowWatcher.Verbose($"FormulaEdit - FormulaBar Unfocus");
                        _formulaEditFocus = FormulaEditFocus.None;
                        UpdateEditState();
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Show:
                        Logger.WindowWatcher.Verbose($"FormulaEdit - FormulaBar Show");
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Hide:
                        Logger.WindowWatcher.Verbose($"FormulaEdit - FormulaBar Hide");
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.LocationChange:
                    if (e.ObjectId == WindowWatcher.WindowChangedEventArgs.ChangeObjectId.Caret)
                    {
                        // We expect this on every text change in newer Excel versions
                        Debug.Print($"-#-#-#- Text Changed ... ");
                        UpdateEditStateDelayed();
                    }
                    else
                    {
                        Debug.Print($"-#-#-#- Unexpected WindowsChanged object id: {e.ObjectId}");
                    }
                    break;
                default:
                    //throw new ArgumentOutOfRangeException("Unexpected Window Change Type", "e.Type");
                    Logger.WindowWatcher.Verbose($"FormulaEdit - Unexpected Window Change Type: {e.Type}");
                    break;
            }
        }

        // Runs on the Automation thread
        void _windowWatcher_InCellEditWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs e)
        {
            // Debug.Print($"\r\n%%% InCellEditWindowChanged: {e.ObjectId} - {e.Type}\r\n");
            switch (e.Type)
            {
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Create:
                    if (e.ObjectId == WindowWatcher.WindowChangedEventArgs.ChangeObjectId.Self)
                    {
                        SetEditWindow(e.WindowHandle, ref _hwndInCellEdit);
                        UpdateEditState();
                    }
                    else if (e.ObjectId == WindowWatcher.WindowChangedEventArgs.ChangeObjectId.Caret)
                    {
                        // We expect this on every text change 
                        // NOTE: Not anymore after some Excel / Windows update
                        Debug.Print($"-#-#-#- Text Changed ... ");
                        UpdateEditStateDelayed();
                    }
                    else
                    {
                        Debug.Print($"### Unexpected WindowsChanged object id: {e.ObjectId}");
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Destroy:
                    // We expect this for every text change, but ignore since we react to the Create event
                    //if (_formulaEditFocus == FormulaEditFocus.FormulaBar)
                    //{
                    //    _formulaEditFocus = FormulaEditFocus.None;
                    //    UpdateEditState();
                    //}
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Focus:
                    if (_formulaEditFocus != FormulaEditFocus.InCellEdit)
                    {
                        if (e.WindowHandle != _hwndInCellEdit)
                        {
                            // We never saw the Create...
                            SetEditWindow(e.WindowHandle, ref _hwndInCellEdit);
                        }

                        Logger.WindowWatcher.Verbose($"FormulaEdit - InCellEdit Focus");

                        _formulaEditFocus = FormulaEditFocus.InCellEdit;
                        UpdateEditState();
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Unfocus:
                    if (_formulaEditFocus == FormulaEditFocus.InCellEdit)
                    {
                        Logger.WindowWatcher.Verbose($"FormulaEdit - InCellEdit Unfocus");
                        _formulaEditFocus = FormulaEditFocus.None;
                        UpdateEditState();
                    }
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Show:
                    Logger.WindowWatcher.Verbose($"FormulaEdit - InCellEdit Show");
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.Hide:
                    Logger.WindowWatcher.Verbose($"FormulaEdit - InCellEdit Hide");
                    break;
                case WindowWatcher.WindowChangedEventArgs.ChangeType.LocationChange:
                    if (e.ObjectId == WindowWatcher.WindowChangedEventArgs.ChangeObjectId.Caret)
                    {
                        // We expect this on every text change in newer Excel versions
                        Debug.Print($"-#-#-#- Text Changed ... ");
                        UpdateEditStateDelayed();
                    }
                    else
                    {
                        Debug.Print($"-#-#-#- Unexpected WindowsChanged object id: {e.ObjectId}");
                    }
                    break;
                default:
                    //throw new ArgumentOutOfRangeException("Unexpected Window Change Type", "e.Type");
                    Logger.WindowWatcher.Verbose($"FormulaEdit - Unexpected Window Change Type: {e.Type}");
                    break;
            }
        }

        // Runs on the automation thread
        void SetEditWindow(IntPtr newWindowHandle, ref IntPtr hwnd)
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

                if (_formulaEditFocus != FormulaEditFocus.None)
                {
                    try
                    {
                        UninstallLocationMonitor();
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
        }

        // Runs on our Automation thread
        void InstallLocationMonitor(IntPtr hWnd)
        {
            UninstallLocationMonitor();
            _windowLocationWatcher = new WindowLocationWatcher(hWnd, _syncContextAuto, _syncContextMain);
            _windowLocationWatcher.LocationChanged += _windowLocationWatcher_LocationChanged;
        }

        // Runs on our Automation thread
        void UninstallLocationMonitor()
        {
            WindowLocationWatcher tempWatcher = Interlocked.Exchange(ref _windowLocationWatcher, null);
            if (tempWatcher != null)
            {
                _syncContextMain.Post(disp => ((IDisposable)disp).Dispose(), tempWatcher);
            }
        }

        // Runs on our Automation thread
        void _windowLocationWatcher_LocationChanged(object sender, EventArgs e)
        {
            UpdateEditState(moveOnly: true);
            _windowWatcher.OnFormulaEditLocationChanged();
        }

        // Runs on an Automation event thread
        // CONSIDER: With WinEvents we could get notified from main thread ... ?
        //void TextChanged(object sender, AutomationEventArgs e)
        //{
        //    // Debug.Print($">>>> FormulaEditWatcher.TextChanged on thread {Thread.CurrentThread.ManagedThreadId}");
        //    Logger.WindowWatcher.Verbose($"FormulaEdit - Text changed in {(sender.Equals(_formulaBar) ? "FormulaBar" : (sender.Equals(_inCellEdit) ? "InCellEdit" : "UNKNOWN"))}");
        //    UpdateFormula(textChangedOnly: true);
        //}

        // TODO: Get rid of this somehow - added to make the mouse clicks in the in-cell editing work, by delaying the call to the PenHelper
        void UpdateEditStateDelayed()
        {
            _syncContextAuto.Post(_ =>
           {
               Thread.Sleep(100);
               UpdateEditState();
           }, null);
        }

        // Switches to our Automation thread, updates current state and raises StateChanged event
        void UpdateEditState(bool moveOnly = false)
        {
            Logger.WindowWatcher.Verbose($"> FormulaEdit UpdateEditState - Thread {Thread.CurrentThread.ManagedThreadId}");
            Logger.WindowWatcher.Verbose($"FormulaEdit UpdateEditState - Focus: {_formulaEditFocus} Window: {(_formulaEditFocus == FormulaEditFocus.FormulaBar ? _hwndFormulaBar : _hwndInCellEdit)}");
            
            IntPtr hwnd = IntPtr.Zero;
            bool prefixChanged = false;
            if (_formulaEditFocus == FormulaEditFocus.FormulaBar)
            {
                hwnd = _hwndFormulaBar;
            }
            else if (_formulaEditFocus == FormulaEditFocus.InCellEdit)
            {
                hwnd = _hwndInCellEdit;
            }
            else
            {
                // Neither have the focus, so we don't update anything
                Logger.WindowWatcher.Verbose("FormulaEdit UpdateEditState End formula editing");
                CurrentPrefix = null;
                if (IsEditingFormula)
                    UninstallLocationMonitor();
                IsEditingFormula = false;
                prefixChanged = true;
                // Debug.Print("#### FormulaEditWatcher - No Window " + Environment.StackTrace);
            }

            if (hwnd != IntPtr.Zero)
            {
                EditWindowBounds = Win32Helper.GetWindowBounds(hwnd);

                if (!IsEditingFormula)
                {
                    IntPtr hwndTopLevel = Win32Helper.GetRootAncestor(hwnd);
                    InstallLocationMonitor(hwndTopLevel);
                    IsEditingFormula = true;
                }

                var newPrefix = XlCall.GetFormulaEditPrefix();  // What thread do we have to use here ...?
                if (CurrentPrefix != newPrefix)
                {
                    CurrentPrefix = newPrefix;
                    prefixChanged = true;
                }
                Logger.WindowWatcher.Verbose($"FormulaEdit UpdateEditState Formula editing: CurrentPrefix {CurrentPrefix}, EditWindowBounds: {EditWindowBounds}");
            }

            // TODO: Smarter (or more direct) notification...?
            if (moveOnly && !prefixChanged)
            {
                StateChanged?.Invoke(this, new StateChangeEventArgs(StateChangeType.Move));
            }
            else
            {
                OnStateChanged(StateChangeType.Multiple);
            }
        }

        // We ensure that our event is raised on the Automation thread .. (Eases concurrency issues in handling it, though it will get passed on to the main thread...)
        void OnStateChanged(StateChangeType stateChangeType)
        {
            StateChangeEventArgs stateChangedArgs;
            switch (stateChangeType)
            {
                case StateChangeType.Multiple:
                    stateChangedArgs = s_stateChangeMultiple;
                    break;
                case StateChangeType.Move:
                    stateChangedArgs = s_stateChangeMove;
                    break;
                case StateChangeType.TextChange:
                    stateChangedArgs = s_stateChangeTextChange;
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(stateChangeType));

            }
            StateChanged?.Invoke(this, stateChangedArgs);
        }

        // Runs on the main thread
        public void Dispose()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);

            Logger.WindowWatcher.Verbose("FormulaEdit Dispose Begin");
            _windowWatcher.FormulaBarWindowChanged -= _windowWatcher_FormulaBarWindowChanged;
            _windowWatcher.InCellEditWindowChanged -= _windowWatcher_InCellEditWindowChanged;

            // Can't call UninstallLocationMonitor - we might be shutting down on the main thread, and don't want to post
            WindowLocationWatcher tempWatcher = Interlocked.Exchange(ref _windowLocationWatcher, null);
            if (tempWatcher != null)
            {
                tempWatcher.Dispose();
            }
            Logger.WindowWatcher.Verbose("FormulaEdit Dispose End");
        }
    }
}
