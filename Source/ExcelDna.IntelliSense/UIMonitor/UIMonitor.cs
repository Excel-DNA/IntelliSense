using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using System.Windows.Automation;

namespace ExcelDna.IntelliSense
{

    // These are immutable representations of the state (reflecting only our interests)
    // We make a fresh a simplified state representation, so that we can make a matching state update representation.
    abstract class UIState
    {
        public static UIState ReadyState = new Ready();
        public class Ready : UIState { }
        public class FormulaEdit : UIState
        {
            public IntPtr FormulaEditWindow;    // Window where text entry focus is - either the in-cell edit window, or the formula bar
            public string FormulaPrefix;
            public Rect EditWindowBounds;
            public Rect ExcelTooltipBounds;

            public virtual FormulaEdit WithFormulaEditWindow(IntPtr newFormulaEditWindow)
            {
                return new FormulaEdit
                {
                    FormulaEditWindow = newFormulaEditWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds
                };
            }

            public virtual FormulaEdit WithFormulaPrefix(string newFormulaPrefix)
            {
                return new FormulaEdit
                {
                    FormulaEditWindow = this.FormulaEditWindow,
                    FormulaPrefix = newFormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds
                };
            }

            public virtual FormulaEdit WithBounds(Rect newEditWindowBounds, Rect newExcelTooltipBounds)
            {
                return new FormulaEdit
                {
                    FormulaEditWindow = this.FormulaEditWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = newEditWindowBounds,
                    ExcelTooltipBounds = newExcelTooltipBounds
                };
            }
        }

        public class FunctionList : FormulaEdit // CONSIDER: I'm not sure the hierarchy here has any value.
        {
            public IntPtr FunctionListWindow;
            public Rect FunctionListBounds;
            public string SelectedItemText;
            public Rect SelectedItemBounds;

            public override FormulaEdit WithFormulaEditWindow(IntPtr newFormulaEditWindow)
            {
                return new FunctionList
                {
                    FormulaEditWindow = newFormulaEditWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds,

                    FunctionListWindow = this.FunctionListWindow,
                    SelectedItemText = this.SelectedItemText,
                    SelectedItemBounds = this.SelectedItemBounds,
                    FunctionListBounds = this.FunctionListBounds
                };
            }

            public FunctionList WithFunctionListWindow(IntPtr newFunctionListWindow)
            {
                return new FunctionList
                {
                    FormulaEditWindow = this.FormulaEditWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds,

                    FunctionListWindow = newFunctionListWindow,
                    SelectedItemText = this.SelectedItemText,
                    SelectedItemBounds = this.SelectedItemBounds,
                    FunctionListBounds = this.FunctionListBounds
                };
            }

            public override FormulaEdit WithFormulaPrefix(string newFormulaPrefix)
            {
                return new FunctionList
                {
                    FormulaEditWindow = this.FormulaEditWindow,
                    FormulaPrefix = newFormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds,

                    FunctionListWindow = this.FunctionListWindow,
                    SelectedItemText = this.SelectedItemText,
                    SelectedItemBounds = this.SelectedItemBounds,
                    FunctionListBounds = this.FunctionListBounds
                };
            }

            public override FormulaEdit WithBounds(Rect newEditWindowBounds, Rect newExcelTooltipBounds)
            {
                return new FunctionList
                {
                    FormulaEditWindow = this.FormulaEditWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = newEditWindowBounds,
                    ExcelTooltipBounds = newExcelTooltipBounds,

                    FunctionListWindow = this.FunctionListWindow,
                    SelectedItemText = this.SelectedItemText,
                    SelectedItemBounds = this.SelectedItemBounds,
                    FunctionListBounds = this.FunctionListBounds
                };
            }

            public virtual FunctionList WithSelectedItem(string selectedItemText, Rect selectedItemBounds, Rect listBounds)
            {
                return new FunctionList
                {
                    FormulaEditWindow = this.FormulaEditWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds,

                    FunctionListWindow = this.FunctionListWindow,
                    FunctionListBounds = listBounds,
                    SelectedItemText = selectedItemText,
                    SelectedItemBounds = selectedItemBounds,
                };
            }

            internal FormulaEdit AsFormulaEdit()
            {
                 return new FormulaEdit
                    {
                        FormulaEditWindow = FormulaEditWindow,
                        FormulaPrefix = FormulaPrefix,
                        EditWindowBounds = EditWindowBounds,
                        ExcelTooltipBounds = ExcelTooltipBounds,
                    };
            }
        }

        // Becomes a more general Window ro Dialog to watch (for resize extension)
        public class SelectDataSource : UIState
        {
            public IntPtr SelectDataSourceWindow;
        }

        public override string ToString()
        {
            #if DEBUG
                return $"{GetType().Name}{((this is Ready) ? "" : "\r\n")}{string.Join("\r\n", GetType().GetFields().Select(fld => $"\t{fld.Name}: {fld.GetValue(this)}"))}";
            #else
                return base.ToString();
            #endif
        }

        // TODO: Figure out what to do with this
        public string LogString()
        {
            #if DEBUG
                return $"{GetType().Name}{((this is Ready) ? "" : "\t")}{string.Join("\t", GetType().GetFields().Select(fld => $"\t{fld.Name}: {fld.GetValue(this)}"))}";
            #else
                return ToString();
            #endif
        }

        // This is the universal update check
        // When an event knows exactly what changed (e.g. Text or SelectedItem), it need not call this
        // CONSIDER: How would this look with C# 7.0 patterns?
        public static IEnumerable<UIStateUpdate> GetUpdates(UIState oldState, UIState newState)
        {
            if (oldState is Ready)
            {
                if (newState is Ready)
                {
                    yield break;
                }
                else if (newState is FunctionList)
                {
                    // We generate an intermediate state (!?)
                    var functionList = (FunctionList)newState;
                    var formulaEdit = functionList.AsFormulaEdit();
                    yield return new UIStateUpdate(oldState, formulaEdit, UIStateUpdate.UpdateType.FormulaEditStart);
                    yield return new UIStateUpdate(formulaEdit, newState, UIStateUpdate.UpdateType.FunctionListShow);
                }
                else if (newState is FormulaEdit) // But not FunctionList
                {
                    yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.FormulaEditStart);
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (var update in GetUpdates(oldState, ReadyState))
                        yield return update;
                    yield return new UIStateUpdate(ReadyState, newState, UIStateUpdate.UpdateType.SelectDataSourceShow);
                }
            }
            else if (oldState is FunctionList)  // and thus also FormulaEdit
            {
                if (newState is Ready)
                {
                    // We generate an intermediate state (!?)
                    var formulaEdit = ((FunctionList)oldState).AsFormulaEdit();
                    yield return new UIStateUpdate(oldState, formulaEdit, UIStateUpdate.UpdateType.FunctionListHide);
                    yield return new UIStateUpdate(formulaEdit, newState, UIStateUpdate.UpdateType.FormulaEditEnd);
                }
                else if (newState is FunctionList)
                {
                    var oldStateFL = (FunctionList)oldState;
                    var newStateFL = (FunctionList)newState;
                    foreach (var update in GetUpdates(oldStateFL, newStateFL))
                        yield return update;
                }
                else if (newState is FormulaEdit) // but not FunctionList
                {
                    var oldStateFE = ((FunctionList)oldState).AsFormulaEdit();
                    yield return new UIStateUpdate(oldState, oldStateFE, UIStateUpdate.UpdateType.FunctionListHide);

                    var newStateFE = (FormulaEdit)newState;
                    foreach (var update in GetUpdates(oldStateFE, newStateFE))
                        yield return update;
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (var update in GetUpdates(oldState, ReadyState))
                        yield return update;
                    yield return new UIStateUpdate(ReadyState, newState, UIStateUpdate.UpdateType.SelectDataSourceShow);
                }
            }
            else if (oldState is FormulaEdit)   // but not FunctionList
            {
                if (newState is Ready)
                {
                    yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.FormulaEditEnd);
                }
                else if (newState is FunctionList)
                {
                    // First process any FormulaEdit changes
                    var oldStateFE = (FormulaEdit)oldState;
                    var newStateFE = ((FunctionList)newState).AsFormulaEdit();
                    foreach (var update in GetUpdates(oldStateFE, newStateFE))
                        yield return update;

                    yield return new UIStateUpdate(newStateFE, newState, UIStateUpdate.UpdateType.FunctionListShow);
                }
                else if (newState is FormulaEdit) // but not FunctionList
                {
                    var oldStateFE = (FormulaEdit)oldState;
                    var newStateFE = (FormulaEdit)newState;
                    foreach (var update in GetUpdates(oldStateFE, newStateFE))
                        yield return update;
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (var update in GetUpdates(oldState, ReadyState))
                        yield return update;
                    yield return new UIStateUpdate(ReadyState, newState, UIStateUpdate.UpdateType.SelectDataSourceShow);
                }
            }
            else if (oldState is SelectDataSource)
            {
                if (newState is Ready)
                {
                    yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.SelectDataSourceHide);
                }
                else if (newState is SelectDataSource)
                {
                    var oldStateSDS = (SelectDataSource)oldState;
                    var newStateSDS = (SelectDataSource)newState;
                    if (oldStateSDS.SelectDataSourceWindow != newStateSDS.SelectDataSourceWindow)
                        yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.SelectDataSourceWindowChange);
                }
                else
                {
                    // Go to Ready, then to new state
                    yield return new UIStateUpdate(oldState, ReadyState, UIStateUpdate.UpdateType.SelectDataSourceHide);
                    foreach (var update in GetUpdates(ReadyState, newState))
                        yield return update;
                }
            }
        }

        static IEnumerable<UIStateUpdate> GetUpdates(FormulaEdit oldState, FormulaEdit newState)
        {
            // We generate intermediate states (!?)
            if (oldState.FormulaEditWindow != newState.FormulaEditWindow)
            {
                // Always changes together with Move ...?
                var tempState = oldState.WithFormulaEditWindow(newState.FormulaEditWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditWindowChange);
                oldState = tempState;
            }
            if (oldState.EditWindowBounds != newState.EditWindowBounds ||
                oldState.ExcelTooltipBounds != newState.ExcelTooltipBounds)
            {
                var tempState = oldState.WithBounds(newState.EditWindowBounds, newState.ExcelTooltipBounds);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditMove);
                oldState = tempState;
            }
            if (oldState.FormulaPrefix != newState.FormulaPrefix)
            {
                yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.FormulaEditTextChange);
            }
        }

        static IEnumerable<UIStateUpdate> GetUpdates(FunctionList oldState, FunctionList newState)
        {
            // We generate intermediate states (!?)
            if (oldState.FormulaEditWindow != newState.FormulaEditWindow)
            {
                // Always changes together with Move ...?
                var tempState = oldState.WithFormulaEditWindow(newState.FormulaEditWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditWindowChange);
                oldState = (FunctionList)tempState;
            }
            if (oldState.FunctionListWindow != newState.FunctionListWindow)
            {
                Debug.Print(">>>>> Unexpected FunctionListWindowChange");  // Should never change???
                var tempState = oldState.WithFunctionListWindow(newState.FunctionListWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FunctionListWindowChange);
                oldState = tempState;
            }
            if (oldState.EditWindowBounds != newState.EditWindowBounds ||
                oldState.ExcelTooltipBounds != newState.ExcelTooltipBounds)
            {
                var tempState = oldState.WithBounds(newState.EditWindowBounds, newState.ExcelTooltipBounds);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditMove);
                oldState = (FunctionList)tempState;
            }
            if (oldState.FormulaPrefix != newState.FormulaPrefix)
            {
                var tempState = oldState.WithFormulaPrefix(newState.FormulaPrefix);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditTextChange);
                oldState = (FunctionList)tempState;
            }
            if (oldState.SelectedItemText != newState.SelectedItemText ||
                oldState.SelectedItemBounds != newState.SelectedItemBounds ||
                oldState.FunctionListBounds != newState.FunctionListBounds)
            {
                yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.FunctionListSelectedItemChange);
            }
        }
    }

    class UIStateUpdate : EventArgs
    {
        // We want to order and nest the updates to make them easy to respond to.
        // This means we have XXXStart, then stuff on the inside, then XXXEnd, always with correct nesting
        public enum UpdateType
        {
            FormulaEditStart,
                // These three updates can happen while FunctionList is shown
                FormulaEditMove,    // Includes moving between in-cell editing and the formula text box.
                FormulaEditWindowChange, // Includes moving between in-cell editing and the formula text box.
                FormulaEditTextChange,

                FunctionListShow,
                    FunctionListMove,
                    FunctionListSelectedItemChange,
                    FunctionListWindowChange,
                FunctionListHide,

            FormulaEditEnd,

            SelectDataSourceShow,
                SelectDataSourceWindowChange,
            SelectDataSourceHide
        }
        public UIState OldState { get; }
        public UIState NewState { get; }
        public UpdateType Update { get; }
        public bool IsEnabled { get; private set; }  // Should this update be raised on the main thread - allows preview event to filter out some events

        public UIStateUpdate(UIState oldState, UIState newState, UpdateType update)
        {
            OldState = oldState;
            NewState = newState;
            Update = update;
            IsEnabled = false;
        }

        // Call this to allow the update event (on the main thread) to be raised
        public void EnableUpdateEvent()
        {
            IsEnabled = true;
        }

        public override string ToString()
        {
            return $"{Update}: {OldState} -> {NewState}";
        }

        public string LogString()
        {
            return $"({Update.ToString()}): [{OldState.LogString()}] -> [{NewState.LogString()}]";
        }
    }

    // Combines all the information received from the UIAutomation and WinEvents watchers,
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

        public UIState CurrentState = UIState.ReadyState;
        public EventHandler<UIStateUpdate> StateUpdatePreview;   // Always called on the automation thread
        public EventHandler<UIStateUpdate> StateUpdate;       // Always posted to the main thread

        public UIMonitor(SynchronizationContext syncContextMain)
        {
            _syncContextMain = syncContextMain;

            // Make a separate thread and set to MTA, according to: https://msdn.microsoft.com/en-us/library/windows/desktop/ee671692%28v=vs.85%29.aspx
            // This thread will be used for UI Automation calls, particularly adding and removing event handlers.
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
            _windowWatcher = new WindowWatcher(_syncContextAuto);
            _formulaEditWatcher = new FormulaEditWatcher(_windowWatcher, _syncContextAuto);
            _popupListWatcher = new PopupListWatcher(_windowWatcher, _syncContextAuto);

            // These are the events we're interested in for showing, hiding and updating the IntelliSense forms
  //          _windowWatcher.MainWindowChanged += _windowWatcher_MainWindowChanged;
            // _windowWatcher.SelectDataSourceWindowChanged += _windowWatcher_SelectDataSourceWindowChanged;
            _popupListWatcher.SelectedItemChanged += _popupListWatcher_SelectedItemChanged;
            _formulaEditWatcher.StateChanged += _formulaEditWatcher_StateChanged;

            _windowWatcher.TryInitialize();

            _syncContextAuto.RunOnCurrentThread();
        }

#region Receive watcher events (all on the automation thread), and process
        //// Runs on our automation thread
        //void _windowWatcher_MainWindowChanged(object sender, EventArgs args)
        //{
        //    Logger.Monitor.Verbose("!> MainWindowChanged");
        //    Logger.Monitor.Verbose("!> " + ReadCurrentState().ToString());
        //    OnStateChanged();
        //}

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
            Logger.Monitor.Verbose($"!> FormulaEdit StateChanged ({args.StateChangeType})");
            //Logger.Monitor.Verbose("!> " + ReadCurrentState().ToString());

            if (args.StateChangeType == FormulaEditWatcher.StateChangeType.TextChange &&
                CurrentState is UIState.FormulaEdit)
            {
                var newState = ((UIState.FormulaEdit)CurrentState).WithFormulaPrefix(_formulaEditWatcher.CurrentPrefix);
                OnStateChanged(newState);
            }
            else
            {
                OnStateChanged();
            }
        }

        // Runs on our automation thread
        UIState ReadCurrentState()
        {
            // The sequence here is important, since we use the hierarchy of UIState classes.
            // if (_selectDataSourceWatcher.
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
                    FormulaPrefix = _formulaEditWatcher.CurrentPrefix ?? "", // TODO: Deal with nulls here... (we're not in FormulaEdit state anymore)
                };
            }
            if (_formulaEditWatcher.IsEditingFormula)
            {
                return new UIState.FormulaEdit
                {
                    FormulaEditWindow = _formulaEditWatcher.FormulaEditWindow,
                    EditWindowBounds = _formulaEditWatcher.EditWindowBounds,
                    FormulaPrefix = _formulaEditWatcher.CurrentPrefix ?? "", // TODO: Deal with nulls here... (we're not in FormulaEdit state anymore)
                };
            }
            //if (_windowWatcher.IsSelectDataSourceWindowVisible)
            //{
            //    return new UIState.SelectDataSource
            //    {
            //        SelectDataSourceWindow = _windowWatcher.SelectDataSourceWindow
            //    };
            //}
            return new UIState.Ready();
        }

        // Filter the states changes (on the automation thread) and then raise changes (on the main thread)
        void OnStateChanged(UIState newStateOrNull = null)
        {
            var oldState = CurrentState;
//            if (newStateOrNull == null)
                newStateOrNull = ReadCurrentState();

            // Debug.Print($"NEWSTATE: {newStateOrNull.ToString()}");

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

        public void Dispose()
        {
            Logger.Monitor.Info($"UIMonitor Dispose Begin");

            if (_syncContextAuto == null)
                return;

            // Send is not supported on _syncContextAuto
            _syncContextAuto.Send(delegate
                {
                    // Remove all event handlers ASAP
                    Automation.RemoveAllEventHandlers();
                    if (_windowWatcher != null)
                    {
//                        _windowWatcher.MainWindowChanged -= _windowWatcher_MainWindowChanged;
                        // _windowWatcher.SelectDataSourceWindowChanged -= _windowWatcher_SelectDataSourceWindowChanged;
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

                }, null);

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

