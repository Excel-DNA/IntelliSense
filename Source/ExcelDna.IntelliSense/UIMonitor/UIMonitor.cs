using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading;

namespace ExcelDna.IntelliSense
{

    // Combines all the information received from the WinEvents watchers,
    // and interprets these into a state that reflect our interests.
    // Raises its events on the main thread.
    // TODO: Also should for argument lists, like TRUE / FALSE in VLOOKUP ...
    class UIMonitor : IDisposable
    {
        readonly SynchronizationContext _syncContextMain;       // Updates will be raised on this thread (but filtered on the automation thread)
        SingleThreadSynchronizationContext _syncContextAuto;    // Running on the Automation thread we create here.

        WindowWatcher _windowWatcher;
        FormulaEditWatcher _formulaEditWatcher;
        PopupListWatcher _popupListWatcher;
        ExcelToolTipWatcher _excelToolTipWatcher;
        IntPtr _lastExcelToolTipShown;  // Zero, if it has been hidden again

        public UIState CurrentState = UIState.ReadyState;
        public EventHandler<UIStateUpdate> StateUpdatePreview;   // Always called on the automation thread
        public EventHandler<UIStateUpdate> StateUpdate;       // Always posted to the main thread

        public UIMonitor(SynchronizationContext syncContextMain)
        {
            _syncContextMain = syncContextMain;

            // Make a separate thread and set to MTA, according to: https://msdn.microsoft.com/en-us/library/windows/desktop/ee671692%28v=vs.85%29.aspx
            // This thread was initially intended for UI Automation calls, particularly adding and removing event handlers.
            var threadAuto = new Thread(RunUIAutomation);
            threadAuto.SetApartmentState(ApartmentState.MTA);
            threadAuto.Start();
        }

        // This runs on the new thread we've created to do all the Automation stuff (_threadAuto)
        // It returns only after when the SyncContext.Complete() has been called (from the IntelliSenseDisplay.Dispose() below)
        void RunUIAutomation()
        {
            _syncContextAuto = new SingleThreadSynchronizationContext();

            Logger.Monitor.Verbose("UIMonitor.RunUIAutomation installing watchers");

            // Create and hook together the various watchers
            _windowWatcher = new WindowWatcher(_syncContextAuto, _syncContextMain);
            _formulaEditWatcher = new FormulaEditWatcher(_windowWatcher, _syncContextAuto, _syncContextMain);
            _popupListWatcher = new PopupListWatcher(_windowWatcher, _syncContextAuto, _syncContextMain);
            _excelToolTipWatcher = new ExcelToolTipWatcher(_windowWatcher, _syncContextAuto);

            // These are the events we're interested in for showing, hiding and updating the IntelliSense forms
            _popupListWatcher.SelectedItemChanged += _popupListWatcher_SelectedItemChanged;
            _formulaEditWatcher.StateChanged += _formulaEditWatcher_StateChanged;
            _excelToolTipWatcher.ToolTipChanged += _excelToolTipWatcher_ToolTipChanged;

            _windowWatcher.TryInitialize();

            _syncContextAuto.RunOnCurrentThread();
        }

        #region Receive watcher events (all on the automation thread), and process
        // Runs on our automation thread
        void _windowWatcher_SelectDataSourceWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs args)
        {
            Logger.Monitor.Verbose($"!> SelectDataSourceWindowChanged ({args.Type})");
            OnStateChanged();
        }

        // Runs on our automation thread
        void _popupListWatcher_SelectedItemChanged(object sender, EventArgs args)
        {
            Logger.Monitor.Verbose("!> PopupList SelectedItemChanged");
            //Logger.Monitor.Verbose("!> " + ReadCurrentState().ToString());

            if (CurrentState is UIState.FunctionList && _popupListWatcher.IsVisible)
            {
                var newState = ((UIState.FunctionList)CurrentState).WithSelectedItem(_popupListWatcher.SelectedItemText, 
                                                                                     _popupListWatcher.SelectedItemBounds,
                                                                                     _popupListWatcher.ListBounds);
                OnStateChanged(newState);
            }
            else
            {
                OnStateChanged();
            }
        }

        // Runs on our automation thread
        // CONSIDER: We might want to do some batching of the text edits?
        void _formulaEditWatcher_StateChanged(object sender, FormulaEditWatcher.StateChangeEventArgs args)
        {
            Logger.Monitor.Verbose($"!> FormulaEdit StateChanged ({args.StateChangeType}) - Thread {Thread.CurrentThread.ManagedThreadId}");
            //Logger.Monitor.Verbose("!> " + ReadCurrentState().ToString());

            if (args.StateChangeType == FormulaEditWatcher.StateChangeType.TextChange &&
                CurrentState is UIState.FormulaEdit)
            {
                var newState = ((UIState.FormulaEdit)CurrentState).WithFormulaPrefix(_formulaEditWatcher.CurrentPrefix);
                OnStateChanged(newState);
                return;
            }

            if (args.StateChangeType == FormulaEditWatcher.StateChangeType.Move)
            {
                if (CurrentState is UIState.FunctionList)
                {
                    var newState = ((UIState.FunctionList)CurrentState).WithBounds(_formulaEditWatcher.EditWindowBounds);
                    OnStateChanged(newState);
                    return;
                }
                if (CurrentState is UIState.FormulaEdit)
                {
                    var newState = ((UIState.FormulaEdit)CurrentState).WithBounds(_formulaEditWatcher.EditWindowBounds);
                    OnStateChanged(newState);
                    return;
                }
            }
            OnStateChanged();
        }

        void _excelToolTipWatcher_ToolTipChanged(object sender, ExcelToolTipWatcher.ToolTipChangeEventArgs e)
        {
            // We assume the ExcelToolTip changes happen before the corresponding FunctionList / FormulaEdit changes
            // So we keep track of the last shown one.
            // This allows the FormulaEdit or transitions to pick it up, hopefully not getting confused with the FunctionList tooltip
            // TODO: Check that the tooltip gets cleared from the CurrentState is it is hidden

            Logger.Monitor.Verbose($"!> ExcelToolTip ToolTipChanged Received: {e} with state: {CurrentState.GetType().Name}");

            if (e.ChangeType == ExcelToolTipWatcher.ToolTipChangeType.Show)
            {
                _lastExcelToolTipShown = e.Handle;
            }
            else if (e.ChangeType == ExcelToolTipWatcher.ToolTipChangeType.Hide)
            {
                if (_lastExcelToolTipShown == e.Handle)
                    _lastExcelToolTipShown = _excelToolTipWatcher.GetLastToolTipOrZero();

                var fe = CurrentState as UIState.FormulaEdit;
                if (fe != null)
                {
                    // Either a FormulaEdit or a FunctionList
                    if (fe.ExcelToolTipWindow == e.Handle)
                    {
                        // OK - it's no longer valid
                        // TODO: Manage the state update
                        // This is a kind of pop.... is it right?
                        var newState = fe.WithToolTipWindow(_lastExcelToolTipShown);
                        OnStateChanged(newState);
                    }
                }
            }
        }

        // Runs on our automation thread
        UIState ReadCurrentState()
        {
            // TODO: ExcelToolTipWindow?

            // The sequence here is important, since we use the hierarchy of UIState classes.
            if (_popupListWatcher.SelectedItemText != string.Empty)
            {
                return new UIState.FunctionList
                {
                    FormulaEditWindow = _formulaEditWatcher.FormulaEditWindow,
                    FunctionListWindow = _popupListWatcher.PopupListHandle,
                    SelectedItemText = _popupListWatcher.SelectedItemText,
                    SelectedItemBounds = _popupListWatcher.SelectedItemBounds,
                    FunctionListBounds = _popupListWatcher.ListBounds,
                    EditWindowBounds = _formulaEditWatcher.EditWindowBounds,
                    FormulaPrefix = _formulaEditWatcher.CurrentPrefix ?? "",
                    ExcelToolTipWindow = _lastExcelToolTipShown // We also keep track here, since we'll by inferring the UIState change list using this too
                };
            }
            if (_formulaEditWatcher.IsEditingFormula)
            {
                return new UIState.FormulaEdit
                {
                    FormulaEditWindow = _formulaEditWatcher.FormulaEditWindow,
                    EditWindowBounds = _formulaEditWatcher.EditWindowBounds,
                    FormulaPrefix = _formulaEditWatcher.CurrentPrefix ?? "",
                    ExcelToolTipWindow = _lastExcelToolTipShown
                };
            }
            //if (_windowWatcher.IsSelectDataSourceWindowVisible)
            //{
            //    return new UIState.SelectDataSource
            //    {
            //        SelectDataSourceWindow = _windowWatcher.SelectDataSourceWindow
            //    };
            //}
            return UIState.ReadyState;
        }

        // Filter the states changes (on the automation thread) and then raise changes (on the main thread)
        void OnStateChanged(UIState newStateOrNull = null)
        {
            var oldState = CurrentState;
            if (newStateOrNull == null)
                newStateOrNull = ReadCurrentState();

            // Debug.Print($"NEWSTATE: {newStateOrNull.ToString()} // {(newStateOrNull is UIState.Ready ? Environment.StackTrace : string.Empty)}");

            CurrentState = newStateOrNull;

            var updates = new List<UIStateUpdate>();    // TODO: Improve perf for common single-update case
            foreach (var update in UIState.GetUpdates(oldState, CurrentState))
            {
                Logger.Monitor.Verbose($">> {update.LogString()}");
                // First we raise the preview event on this thread
                // allowing listeners to enable the main-thread event for this update
                StateUpdatePreview?.Invoke(this, update);
                if (update.IsEnabled)
                    updates.Add(update);
            }
            if (updates.Count > 0)
                _syncContextMain.Post(RaiseStateUpdates, updates);
        }

        // Perf experiment
        //void OnStateChanged(UIState newStateOrNull = null)
        //{
        //    var oldState = CurrentState;
        //    if (newStateOrNull == null)
        //        newStateOrNull = ReadCurrentState();
        //    CurrentState = newStateOrNull;
        //    // TODO: Performance: Can we get rid of this list...?
        //    var updateEnumerator = UIState.GetUpdates(oldState, CurrentState).GetEnumerator();
        //    if (!updateEnumerator.MoveNext())
        //        return;
        //    var firstUpdate = updateEnumerator.Current;
        //    if (!updateEnumerator.MoveNext())
        //    {
        //        // Only had the one item
        //        if (StateUpdateFilter(firstUpdate))
        //        {
        //            _syncContextMain.Post(SingleStateChange, firstUpdate);
        //        }
        //        return;

        //    }
        //    var updates = new List<UIStateUpdate>();
        //    if (StateUpdateFilter(firstUpdate))
        //        updates.Add(firstUpdate);
        //    do
        //    {
        //        var update = updateEnumerator.Current; 
        //        if (StateUpdateFilter(update))
        //        updates.Add(update);
        //    } while (updateEnumerator.MoveNext());

        //    Debug.Print($"MULTIPLE STATE UPDATES: {updates.Count}");
        //    _syncContextMain.Post(MultipleStateChanges, updates);
        //}

        // Runs on the main thread
        void RaiseStateUpdates(object updates)
        {
            foreach (var update in (List<UIStateUpdate>)updates)
            {
                StateUpdate?.Invoke(this, update);
            }
        }

#endregion

        // Must run on the main thread
        public void Dispose()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.Monitor.Info($"UIMonitor Dispose Begin");

            // Remove all event handlers ASAP
            // Since we are running on the main thread, we call Dispose directly
            // (might not be in a context where we can post to or wait for main thread sync context)
            if (_windowWatcher != null)
            {
                _windowWatcher.Dispose();
                _windowWatcher = null;
            }
            if (_formulaEditWatcher != null)
            {
                _formulaEditWatcher.StateChanged -= _formulaEditWatcher_StateChanged;
                _formulaEditWatcher.Dispose();
                _formulaEditWatcher = null;
            }
            if (_popupListWatcher != null)
            {
                _popupListWatcher.SelectedItemChanged -= _popupListWatcher_SelectedItemChanged;
                _popupListWatcher.Dispose();
                _popupListWatcher = null;
            }

            if (_syncContextAuto == null)
            {
                Debug.Fail("Unexpected");
                return;
            }

            // Let the above delegate and nested calls run, then clean up.
            // (not sure it makes a difference anymore...)
            _syncContextAuto.Post(delegate
            {
                _syncContextAuto.Complete();
                _syncContextAuto = null;
            }, null);
            Logger.Monitor.Info($"UIMonitor Dispose End");
        }
    }
}

