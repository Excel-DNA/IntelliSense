using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;

namespace ExcelDna.IntelliSense
{

    // These are immutable representations of the state (reflecting only our interests)
    // We make a fresh a simplified state representation, so that we can make a matching state update representation.
    public abstract class UIState
    {
        public static UIState ReadyState = new Ready();
        public class Ready : UIState { }
        public class FormulaEdit : UIState
        {
            public IntPtr MainWindow;
            public string FormulaPrefix;
            public Rect EditWindowBounds;
            public Rect ExcelTooltipBounds;

            public virtual FormulaEdit WithMainWindow(IntPtr newMainWindow)
            {
                return new FormulaEdit
                {
                    MainWindow = newMainWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds
                };
            }

            public virtual FormulaEdit WithFormulaPrefix(string newFormulaPrefix)
            {
                return new FormulaEdit
                {
                    MainWindow = this.MainWindow,
                    FormulaPrefix = newFormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds
                };
            }

            public virtual FormulaEdit WithBounds(Rect newEditWindowBounds, Rect newExcelTooltipBounds)
            {
                return new FormulaEdit
                {
                    MainWindow = this.MainWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = newEditWindowBounds,
                    ExcelTooltipBounds = newExcelTooltipBounds
                };
            }
        }

        public class FunctionList : FormulaEdit // CONSIDER: I'm not sure the hierarchy here has any value.
        {
            public string SelectedItemText;
            public Rect SelectedItemBounds;

            public override FormulaEdit WithMainWindow(IntPtr newMainWindow)
            {
                return new FunctionList
                {
                    MainWindow = newMainWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds,
                    SelectedItemText = this.SelectedItemText,
                    SelectedItemBounds = this.SelectedItemBounds
                };
            }

            public override FormulaEdit WithFormulaPrefix(string newFormulaPrefix)
            {
                return new FunctionList
                {
                    MainWindow = this.MainWindow,
                    FormulaPrefix = newFormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds,
                    SelectedItemText = this.SelectedItemText,
                    SelectedItemBounds = this.SelectedItemBounds
                };
            }

            public override FormulaEdit WithBounds(Rect newEditWindowBounds, Rect newExcelTooltipBounds)
            {
                return new FunctionList
                {
                    MainWindow = this.MainWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = newEditWindowBounds,
                    ExcelTooltipBounds = newExcelTooltipBounds,
                    SelectedItemText = this.SelectedItemText,
                    SelectedItemBounds = this.SelectedItemBounds
                };
            }

            public virtual FunctionList WithSelectedItem(string selectedItemText, Rect selectedItemBounds)
            {
                return new FunctionList
                {
                    MainWindow = this.MainWindow,
                    FormulaPrefix = this.FormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds,
                    SelectedItemText = selectedItemText,
                    SelectedItemBounds = selectedItemBounds
                };
            }

            internal FormulaEdit AsFormulaEdit()
            {
                 return new FormulaEdit
                    {
                        EditWindowBounds = EditWindowBounds,
                        ExcelTooltipBounds = ExcelTooltipBounds,
                        FormulaPrefix = FormulaPrefix,
                        MainWindow = MainWindow
                    };
            }
        }
        public class SelectDataSource : UIState
        {
            public IntPtr MainWindow;
            public IntPtr SelectDataSourceWindow;
        }

        public override string ToString()
        {
            return $"{GetType().Name}{((this is Ready) ? "" : "\r\n")}{string.Join("\r\n", GetType().GetFields().Select(fld => $"\t{fld.Name}: {fld.GetValue(this)}"))}";
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
                    if (oldStateSDS.MainWindow != newStateSDS.MainWindow)
                        yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.SelectDataSourceMainWindowChange);
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
            if (oldState.MainWindow != newState.MainWindow)
            {
                var tempState = oldState.WithMainWindow(newState.MainWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditMainWindowChange);
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
            if (oldState.MainWindow != newState.MainWindow)
            {
                var tempState = oldState.WithMainWindow(newState.MainWindow);
                yield return new UIStateUpdate(oldState, tempState, UIStateUpdate.UpdateType.FormulaEditMainWindowChange);
                oldState = (FunctionList)tempState;
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
                oldState.SelectedItemBounds != newState.SelectedItemBounds)
            {
                yield return new UIStateUpdate(oldState, newState, UIStateUpdate.UpdateType.FunctionListSelectedItemChange);
            }
        }

    }

    public class UIStateUpdate : EventArgs
    {
        // We want to order and nest the updates to make them easy to respond to.
        // This means we have XXXStart, then stuff on the inside, then XXXEnd, always with correct nesting
        public enum UpdateType
        {
            FormulaEditStart,
                // These two can happen while FunctionList is shown
                FormulaEditMove,    // Includes moving between in-cell editing and the formula text box.
                FormulaEditMainWindowChange,
                FormulaEditTextChange,

                FunctionListShow,
                    FunctionListMove,
                    FunctionListSelectedItemChange,
                FunctionListHide,

            FormulaEditEnd,

            SelectDataSourceShow,
                SelectDataSourceMainWindowChange,
            SelectDataSourceHide
        }
        public UIState OldState;
        public UIState NewState;
        public UpdateType Update;

        public UIStateUpdate(UIState oldState, UIState newState, UpdateType update)
        {
            OldState = oldState;
            NewState = newState;
            Update = update;
        }
    }

    // Combines all the information received from the UIAutomation and WinEvents watchers,
    // and interprets these into a state that reflect our interests.
    // Raises its events on the main thread.
    // TODO: Also should for argument lists, like TRUE / FALSE in VLOOKUP ...
    class UIMonitor : IDisposable
    {
        readonly SynchronizationContext _syncContextMain;
        SingleThreadSynchronizationContext _syncContextAuto;    // Running on the Automation thread we create here.

        WindowWatcher _windowWatcher;
        FormulaEditWatcher _formulaEditWatcher;
        PopupListWatcher _popupListWatcher;

        public UIState CurrentState = UIState.ReadyState;
        public event EventHandler<UIStateUpdate> StateChanged;

        public UIMonitor(SynchronizationContext syncContextMain)
        {
            _syncContextMain = syncContextMain;

            // Make a separate thread and set to MTA, according to: https://msdn.microsoft.com/en-us/library/windows/desktop/ee671692%28v=vs.85%29.aspx
            var threadAuto = new Thread(RunUIAutomation);
            threadAuto.SetApartmentState(ApartmentState.MTA);
            threadAuto.Start();

        }

        // This runs on the new thread we've created to do all the Automation stuff (_threadAuto)
        // It returns only after when the SyncContext.Complete() has been called (from the IntelliSenseDisplay.Dispose() below)
        void RunUIAutomation()
        {
            _syncContextAuto = new SingleThreadSynchronizationContext();

            // Create and hook together the various watchers
            _windowWatcher = new WindowWatcher(_syncContextAuto);
            _formulaEditWatcher = new FormulaEditWatcher(_windowWatcher, _syncContextAuto);
            _popupListWatcher = new PopupListWatcher(_windowWatcher, _syncContextAuto);

            // These are the events we're interested in for showing, hiding and updating the IntelliSense forms
            _windowWatcher.MainWindowChanged += _windowWatcher_MainWindowChanged;
            _windowWatcher.SelectDataSourceWindowChanged += _windowWatcher_SelectDataSourceWindowChanged;
            _popupListWatcher.SelectedItemChanged += _popupListWatcher_SelectedItemChanged;
            _formulaEditWatcher.StateChanged += _formulaEditWatcher_StateChanged;

            _windowWatcher.TryInitialize();

            _syncContextAuto.RunOnCurrentThread();
        }

        #region Receive watcher events (all on the automation thread), and process
        // Runs on our automation thread
        void _windowWatcher_MainWindowChanged(object sender, EventArgs args)
        {
            Logger.Monitor.Verbose("!> MainWindowChanged");
            Logger.Monitor.Verbose("!> " + ReadCurrentState().ToString());
            OnStateChanged();
        }

        // Runs on our automation thread
        void _windowWatcher_SelectDataSourceWindowChanged(object sender, WindowWatcher.WindowChangedEventArgs args)
        {
            Logger.Monitor.Verbose($"!> SelectDataSourceWindowChanged ({args.Type})");
            OnStateChanged();
        }

        void _popupListWatcher_SelectedItemChanged(object sender, EventArgs args)
        {
            Logger.Monitor.Verbose("!> PopupList SelectedItemChanged");
            Logger.Monitor.Verbose("!> " + ReadCurrentState().ToString());

            if (CurrentState is UIState.FunctionList && _popupListWatcher.IsVisible)
            {
                var newState = ((UIState.FunctionList)CurrentState).WithSelectedItem(_popupListWatcher.SelectedItemText, _popupListWatcher.SelectedItemBounds);
                OnStateChanged(newState);
            }
            else
            {
                OnStateChanged();
            }
        }

        // Runs on our automation thread
        void _formulaEditWatcher_StateChanged(object sender, FormulaEditWatcher.StateChangeEventArgs args)
        {
            Logger.Monitor.Verbose($"!> FormulaEdit StateChanged ({args.StateChangeType})");
            Logger.Monitor.Verbose("!> " + ReadCurrentState().ToString());

            if (args.StateChangeType == FormulaEditWatcher.StateChangeType.TextChangedOnly &&
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
                    MainWindow = _windowWatcher.MainWindow,
                    SelectedItemText = _popupListWatcher.SelectedItemText,
                    SelectedItemBounds = _popupListWatcher.SelectedItemBounds,
                    EditWindowBounds = _formulaEditWatcher.EditWindowBounds,
                    FormulaPrefix = _formulaEditWatcher.CurrentPrefix,
                };
            }
            if (_formulaEditWatcher.IsEditingFormula)
            {
                return new UIState.FormulaEdit
                {
                    MainWindow = _windowWatcher.MainWindow,
                    EditWindowBounds = _formulaEditWatcher.EditWindowBounds,
                    FormulaPrefix = _formulaEditWatcher.CurrentPrefix,
                };
            }
            if (_windowWatcher.IsSelectDataSourceWindowVisible)
            {
                return new UIState.SelectDataSource
                {
                    MainWindow = _windowWatcher.MainWindow,
                    SelectDataSourceWindow = _windowWatcher.SelectDataSourceWindow
                };
            }
            return new UIState.Ready();
        }

        // Raises the StateChanged event on our automation thread
        // TODO: We might short-cut the update type too, for the common FormulaEdit and SelectItemChange cases
        void OnStateChanged(UIState newStateOrNull = null)
        {
            var oldState = CurrentState;
            if (newStateOrNull == null)
                newStateOrNull = ReadCurrentState();
            CurrentState = newStateOrNull;
            var updates = UIState.GetUpdates(oldState, CurrentState).ToList();
            if (updates.Count > 1)
            {
                Debug.Print($"MULTIPLE STATE UPDATES: {updates.Count}");
            }
            foreach (var update in updates)
            {
                StateChanged?.Invoke(this, update);
            }
        }
        #endregion

        public void Dispose()
        {

            if (_syncContextAuto == null)
                return;

            // Send is not supported on _syncContextAuto
            _syncContextAuto.Post(delegate
                {
                    if (_windowWatcher != null)
                    {
                        _windowWatcher.MainWindowChanged -= _windowWatcher_MainWindowChanged;
                        _windowWatcher.SelectDataSourceWindowChanged -= _windowWatcher_SelectDataSourceWindowChanged;
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
                    _syncContextAuto.Complete();
                    _syncContextAuto = null;
                }, null);
        }
    }
}

