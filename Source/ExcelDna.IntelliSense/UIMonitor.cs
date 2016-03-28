using System;
using System.Collections.Generic;
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

            public virtual UIState WithFormulaPrefix(string newFormulaPrefix)
            {
                return new FormulaEdit
                {
                    MainWindow = this.MainWindow,
                    FormulaPrefix = newFormulaPrefix,
                    EditWindowBounds = this.EditWindowBounds,
                    ExcelTooltipBounds = this.ExcelTooltipBounds
                };
            }
        }

        public class FunctionList : FormulaEdit // CONSIDER: I'm not sure the hierarchy here has any value.
        {
            public string SelectedItemText;
            public Rect SelectedItemBounds;

            public override UIState WithFormulaPrefix(string newFormulaPrefix)
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

            public virtual UIState WithSelectedItem(string selectedItemText, Rect selectedItemBounds)
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
        public static IEnumerable<UIStateUpdate.UpdateType> GetUpdateTypes(UIState oldState, UIState newState)
        {
            if (oldState is Ready)
            {
                if (newState is Ready)
                {
                    yield break;
                }
                else if (newState is FormulaEdit)
                {
                    yield return UIStateUpdate.UpdateType.FormulaEditStart;
                    if (newState is FunctionList)
                    {
                        yield return UIStateUpdate.UpdateType.FunctionListShow;
                    }
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (var update in GetUpdateTypes(oldState, ReadyState))
                        yield return update;
                    yield return UIStateUpdate.UpdateType.SelectDataSourceShow;
                }
            }
            else if (oldState is FunctionList)  // and thus also FormulaEdit
            {
                if (newState is Ready)
                {
                    yield return UIStateUpdate.UpdateType.FunctionListHide;
                    yield return UIStateUpdate.UpdateType.FormulaEditEnd;
                }
                else if (newState is FunctionList)
                {
                    var oldStateFL = (FunctionList)oldState;
                    var newStateFL = (FunctionList)newState;
                    if (oldStateFL.MainWindow != newStateFL.MainWindow)
                        yield return UIStateUpdate.UpdateType.FormulaEditMainWindowChange;
                    if (oldStateFL.EditWindowBounds != newStateFL.EditWindowBounds ||
                        oldStateFL.ExcelTooltipBounds != newStateFL.ExcelTooltipBounds)
                        yield return UIStateUpdate.UpdateType.FormulaEditMove;
                    if (oldStateFL.FormulaPrefix != newStateFL.FormulaPrefix)
                        yield return UIStateUpdate.UpdateType.FormulaEditTextChange;
                    if (oldStateFL.SelectedItemBounds != newStateFL.SelectedItemBounds ||
                        oldStateFL.SelectedItemText != newStateFL.SelectedItemText)
                        yield return UIStateUpdate.UpdateType.FunctionListSelectedItemChange;
                }
                else if (newState is FormulaEdit)
                {
                    yield return UIStateUpdate.UpdateType.FunctionListHide;
                    var oldStateFE = (FormulaEdit)oldState;
                    var newStateFE = (FormulaEdit)newState;
                    if (oldStateFE.MainWindow != newStateFE.MainWindow)
                        yield return UIStateUpdate.UpdateType.FormulaEditMainWindowChange;
                    if (oldStateFE.EditWindowBounds != newStateFE.EditWindowBounds ||
                        oldStateFE.ExcelTooltipBounds != newStateFE.ExcelTooltipBounds)
                        yield return UIStateUpdate.UpdateType.FormulaEditMove;
                    if (oldStateFE.FormulaPrefix != newStateFE.FormulaPrefix)
                        yield return UIStateUpdate.UpdateType.FormulaEditTextChange;
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (var update in GetUpdateTypes(oldState, ReadyState))
                        yield return update;
                    yield return UIStateUpdate.UpdateType.SelectDataSourceShow;
                }
            }
            else if (oldState is FormulaEdit)   // but not FunctionList
            {
                if (newState is Ready)
                {
                    yield return UIStateUpdate.UpdateType.FormulaEditEnd;
                }
                else if (newState is FunctionList)
                {
                    yield return UIStateUpdate.UpdateType.FunctionListShow;
                    var oldStateFL = (FunctionList)oldState;
                    var newStateFL = (FunctionList)newState;
                    if (oldStateFL.MainWindow != newStateFL.MainWindow)
                        yield return UIStateUpdate.UpdateType.FormulaEditMainWindowChange;
                    if (oldStateFL.EditWindowBounds != newStateFL.EditWindowBounds ||
                        oldStateFL.ExcelTooltipBounds != newStateFL.ExcelTooltipBounds)
                        yield return UIStateUpdate.UpdateType.FormulaEditMove;
                    if (oldStateFL.FormulaPrefix != newStateFL.FormulaPrefix)
                        yield return UIStateUpdate.UpdateType.FormulaEditTextChange;
                }
                else if (newState is FormulaEdit)
                {
                    var oldStateFE = (FormulaEdit)oldState;
                    var newStateFE = (FormulaEdit)newState;
                    if (oldStateFE.MainWindow != newStateFE.MainWindow)
                        yield return UIStateUpdate.UpdateType.FormulaEditMainWindowChange;
                    if (oldStateFE.EditWindowBounds != newStateFE.EditWindowBounds ||
                        oldStateFE.ExcelTooltipBounds != newStateFE.ExcelTooltipBounds)
                        yield return UIStateUpdate.UpdateType.FormulaEditMove;
                    if (oldStateFE.FormulaPrefix != newStateFE.FormulaPrefix)
                        yield return UIStateUpdate.UpdateType.FormulaEditTextChange;
                }
                else if (newState is SelectDataSource)
                {
                    // Go to Ready then to new state
                    foreach (var update in GetUpdateTypes(oldState, ReadyState))
                        yield return update;
                    yield return UIStateUpdate.UpdateType.SelectDataSourceShow;
                }
            }
            else if (oldState is SelectDataSource)
            {
                if (newState is Ready)
                {
                    yield return UIStateUpdate.UpdateType.SelectDataSourceHide;
                }
                else if (newState is SelectDataSource)
                {
                    var oldStateSDS = (SelectDataSource)oldState;
                    var newStateSDS = (SelectDataSource)newState;
                    if (oldStateSDS.MainWindow != newStateSDS.MainWindow)
                        yield return UIStateUpdate.UpdateType.SelectDataSourceMainWindowChange;
                }
                else
                {
                    // Go to Ready, then to new state
                    yield return UIStateUpdate.UpdateType.SelectDataSourceHide;
                    foreach (var update in GetUpdateTypes(ReadyState, newState))
                        yield return update;
                }
            }
        }

    }

    public class UIStateUpdate : EventArgs
    {
        // An Update notification might have multiple UpdateTypes

        // We want to order and nest the updates (and UpdateTypes inside an Update), to make them easy to respond to.
        // This means we have XXXStart, then stuff on the inside, then XXXEnd with correct nesting
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
        public IEnumerable<UIStateUpdate.UpdateType> UpdateTypes;
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
            StateChanged?.Invoke(this,
                new UIStateUpdate
                {
                    OldState = oldState,
                    NewState = CurrentState,
                    UpdateTypes = UIState.GetUpdateTypes(oldState, CurrentState)
                });
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

