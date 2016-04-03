using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense
{
    [Serializable]
    class IntelliSenseFunctionInfo
    {
        [Serializable]
        public class ArgumentInfo
        {
            public string ArgumentName;
            public string Description;
        }
        public string FunctionName;
        public string Description;
        public List<ArgumentInfo> ArgumentList;
        public string SourcePath; // XllPath for .xll, Workbook Name for Workbook
    }

    // CONSIDER: Revisit UI Automation Threading: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671692(v=vs.85).aspx
    //           And this threading sample using tlbimp version of Windows 7 native UIA: http://code.msdn.microsoft.com/Windows-7-UI-Automation-6390614a/sourcecode?fileId=21469&pathId=715901329
    class IntelliSenseDisplay : IDisposable
    {

        // SyncContextMain is running on the main Excel thread (not a 'macro' context, though)
        // Not sure we need this here... (the UIMonitor internally references it, and could raise the update events on the main thread...).
        // CONSIDER: We could send in the two filters for selecteditem and formula change, so that these checks don't run on the main thread...

        SynchronizationContext _syncContextMain;
        readonly UIMonitor _uiMonitor;

        readonly Dictionary<string, IntelliSenseFunctionInfo> _functionInfoMap =
            new Dictionary<string, IntelliSenseFunctionInfo>(StringComparer.CurrentCultureIgnoreCase);

        // Need to make these late ...?
        ToolTipForm _descriptionToolTip;
        ToolTipForm _argumentsToolTip;
        IntPtr _mainWindow;
        
        public IntelliSenseDisplay(SynchronizationContext syncContextMain, UIMonitor uiMonitor)
        {
            // We expect this to be running in a macro context on the main Excel thread (ManagedThreadId = 1).
            #pragma warning disable CS0618 // Type or member is obsolete (GetCurrentThreadId) - But for debugging we want to monitor this anyway
            Debug.Print($"### Thread creating IntelliSenseDisplay: Managed {Thread.CurrentThread.ManagedThreadId}, Native {AppDomain.GetCurrentThreadId()}");
            #pragma warning restore CS0618 // Type or member is obsolete

            _syncContextMain = syncContextMain;
            _uiMonitor = uiMonitor;
            _uiMonitor.StateUpdatePreview += StateUpdatePreview;
            _uiMonitor.StateUpdate += StateUpdate;
        }

        // This runs on the UIMonitor's automation thread
        // Allows us to enable the update to be raised on main thread
        // Might be raised on the main thread even if we don't enable it (other listeners might enable)
        void StateUpdatePreview(object sender,  UIStateUpdate update)
        {
            bool enable;
            if (update.Update == UIStateUpdate.UpdateType.FormulaEditTextChange)
            {
                var fe = (UIState.FormulaEdit)update.NewState;
                enable = ShouldProcessFormulaEditTextChange(fe.FormulaPrefix);
            }
            else if (update.Update == UIStateUpdate.UpdateType.FunctionListSelectedItemChange)
            {
                var fl = (UIState.FunctionList)update.NewState;
                enable = ShouldProcessFunctionListSelectedItemChange(fl.SelectedItemText);
            }
            else
            {
                enable = true; // allow the update event to be raised by default
            }

            if (enable)
                update.EnableUpdateEvent();
        }

        // This runs on the main thread
        void StateUpdate(object sender, UIStateUpdate update)
        {
            Debug.Print($"STATE UPDATE ({update.Update}): {update.OldState} => {update.NewState}");

            switch (update.Update)
            {
                case UIStateUpdate.UpdateType.FormulaEditStart:
                    UpdateMainWindow((update.NewState as UIState.FormulaEdit).MainWindow);
                    FormulaEditStart();
                    break;
                case UIStateUpdate.UpdateType.FormulaEditMove:
                    break;
                case UIStateUpdate.UpdateType.FormulaEditMainWindowChange:
                    UpdateMainWindow((update.NewState as UIState.FormulaEdit).MainWindow);
                    break;
                case UIStateUpdate.UpdateType.FormulaEditTextChange:
                    var fe = (UIState.FormulaEdit)update.NewState;
                    FormulaEditTextChange(fe.FormulaPrefix, fe.EditWindowBounds, fe.ExcelTooltipBounds);
                    break;
                case UIStateUpdate.UpdateType.FunctionListShow:
                    FunctionListShow();
                    var fls = (UIState.FunctionList)update.NewState;
                    FunctionListSelectedItemChange(fls.SelectedItemText, fls.SelectedItemBounds);
                    break;
                case UIStateUpdate.UpdateType.FunctionListMove:
                    break;
                case UIStateUpdate.UpdateType.FunctionListSelectedItemChange:
                    var fl = (UIState.FunctionList)update.NewState;
                    FunctionListSelectedItemChange(fl.SelectedItemText, fl.SelectedItemBounds);
                    break;
                case UIStateUpdate.UpdateType.FunctionListHide:
                    FunctionListHide();
                    break;
                case UIStateUpdate.UpdateType.FormulaEditEnd:
                    FormulaEditEnd();
                    break;
                case UIStateUpdate.UpdateType.SelectDataSourceShow:
                case UIStateUpdate.UpdateType.SelectDataSourceMainWindowChange:
                case UIStateUpdate.UpdateType.SelectDataSourceHide:
                    // We ignore these
                    break;
                default:
                    throw new InvalidOperationException("Unexpected UIStateUpdate");
            }
        }

        // Runs on the main thread
        void UpdateMainWindow(IntPtr mainWindow)
        {
            if (_mainWindow != mainWindow &&
                 (_descriptionToolTip != null ||
                  _argumentsToolTip   != null ))
            {
                if (_descriptionToolTip != null)
                {
                    _descriptionToolTip.Dispose();
                    _descriptionToolTip = new ToolTipForm(_mainWindow);
                }
                if (_argumentsToolTip != null)
                {
                    _argumentsToolTip.Dispose();
                    _argumentsToolTip = new ToolTipForm(_mainWindow);
                }
                _mainWindow = mainWindow;
            }
        }

        // Runs on the main thread
        void FunctionListShow()
        {
            if (_descriptionToolTip == null)
                _descriptionToolTip = new ToolTipForm(_mainWindow);
        }

        // Runs on the main thread
        void FunctionListHide()
        {
            _descriptionToolTip.Hide();
        }

        // Runs on the main thread
        void FormulaEditStart()
        {
            if (_argumentsToolTip == null)
                _argumentsToolTip = new ToolTipForm(_mainWindow);
        }

        // Runs on the main thread
        void FormulaEditTextChange(string formulaPrefix, Rect editWindowBounds, Rect excelTooltipBounds)
        {
            Debug.Print($"^^^ FormulaEditStateChanged. CurrentPrefix: {formulaPrefix}, Thread {Thread.CurrentThread.ManagedThreadId}");
            var match = Regex.Match(formulaPrefix, @"^=(?<functionName>\w*)\(");
            if (match.Success)
            {
                string functionName = match.Groups["functionName"].Value;

                IntelliSenseFunctionInfo functionInfo;
                if (_functionInfoMap.TryGetValue(functionName, out functionInfo))
                {
                    // TODO: Fix this: Need to consider subformulae
                    int currentArgIndex = formulaPrefix.Count(c => c == ',');
                    _argumentsToolTip.ShowToolTip(
                        GetFunctionIntelliSense(functionInfo, currentArgIndex),
                        (int)editWindowBounds.Left, (int)editWindowBounds.Bottom + 5);
                    return;
                }
            }

            // All other paths, we just clear the box
            _argumentsToolTip.Hide();
        }

        // Runs on the main thread
        void FormulaEditEnd()
        {
            _argumentsToolTip.Hide();
        }

        // Runs on the UIMonitor's automation thread - return true if we might want to process
        bool ShouldProcessFunctionListSelectedItemChange(string selectedItemText)
        {
            if (_descriptionToolTip?.Visible == true)
                return true;
            
            return _functionInfoMap.ContainsKey(selectedItemText);
        }

        // Runs on the UIMonitor's automation thread - return true if we might want to process
        bool ShouldProcessFormulaEditTextChange(string formulaPrefix)
        {
            // CAREFUL: Because of threading, this might run before FormulaEditStart!

            if (_argumentsToolTip?.Visible == true)
                return true;

            // TODO: Consolidate the check here with the FormulaMonitor
            var match = Regex.Match(formulaPrefix, @"^=(?<functionName>\w*)\(");
            if (match.Success)
            {
                string functionName = match.Groups["functionName"].Value;
                return _functionInfoMap.ContainsKey(functionName);
            }
            // Not interested...
            Debug.Print($"Not processing formula {formulaPrefix}");
            return false;
        }
        
        // Runs on the main thread
        void FunctionListSelectedItemChange(string selectedItemText, Rect selectedItemBounds)
        {
            Debug.Print($"IntelliSenseDisplay - PopupListSelectedItemChanged - New text - {selectedItemText}, Thread {Thread.CurrentThread.ManagedThreadId}");

            IntelliSenseFunctionInfo functionInfo;
            if (_functionInfoMap.TryGetValue(selectedItemText, out functionInfo))
            {
                // It's ours!
                _descriptionToolTip.ShowToolTip(
                    text: new FormattedText { GetFunctionDescription(functionInfo) },
                    left: (int)selectedItemBounds.Right + 25,
                    top: (int)selectedItemBounds.Top);
            }
            else
            {
                FunctionListHide();
            }
        }

        //// Runs on the main thread
        //// TODO: Need better formula parsing story here
        //// Here are some ideas: http://fastexcel.wordpress.com/2013/10/27/parsing-functions-from-excel-formulas-using-vba-is-mid-or-a-byte-array-the-best-method/
        //void FormulaEditStateChanged(object stateChangeTypeObj)
        //{
        //    var stateChangeType = (FormulaEditWatcher.StateChangeType)stateChangeTypeObj;
        //    // Check for watcher already disposed 
        //    // CONSIDER: How to manage threading with disposal...?
        //    if (_formulaEditWatcher == null)
        //        return;

        //    if (stateChangeType == FormulaEditWatcher.StateChangeType.Move && _argumentsToolTip != null)
        //    {
        //        _argumentsToolTip.MoveToolTip(
        //            (int)_formulaEditWatcher.EditWindowBounds.Left, (int)_formulaEditWatcher.EditWindowBounds.Bottom + 5);
        //        return;
        //    }

        //    Debug.Print($"^^^ FormulaEditStateChanged. CurrentPrefix: {_formulaEditWatcher.CurrentPrefix}, Thread {Thread.CurrentThread.ManagedThreadId}");
        //    if (_formulaEditWatcher.IsEditingFormula && _formulaEditWatcher.CurrentPrefix != null)
        //    {
        //        string prefix = _formulaEditWatcher.CurrentPrefix;
        //        var match = Regex.Match(prefix, @"^=(?<functionName>\w*)\(");
        //        if (match.Success)
        //        {
        //            string functionName = match.Groups["functionName"].Value;

        //            IntelliSenseFunctionInfo functionInfo;
        //            if (_functionInfoMap.TryGetValue(functionName, out functionInfo))
        //            {
        //                // It's ours!
        //                if (_argumentsToolTip == null)
        //                {
        //                    _argumentsToolTip = new ToolTipForm(_windowWatcher.MainWindow);
        //                }

        //                // TODO: Fix this: Need to consider subformulae
        //                int currentArgIndex = _formulaEditWatcher.CurrentPrefix.Count(c => c == ',');
        //                _argumentsToolTip.ShowToolTip(
        //                    GetFunctionIntelliSense(functionInfo, currentArgIndex),
        //                    (int)_formulaEditWatcher.EditWindowBounds.Left, (int)_formulaEditWatcher.EditWindowBounds.Bottom + 5);
        //                return;
        //            }
        //        }
        //    }

        //    // All other paths, we just clear the box
        //    if (_argumentsToolTip != null)
        //        _argumentsToolTip.Hide();
        //}


        // TODO: Performance / efficiency - cache these somehow
        // TODO: Probably not a good place for LINQ !?
        static readonly string[] s_newLineStringArray = new string[] { Environment.NewLine };
        IEnumerable<TextLine> GetFunctionDescription(IntelliSenseFunctionInfo functionInfo)
        {
            return 
                functionInfo.Description
                .Split(s_newLineStringArray, StringSplitOptions.None)
                .Select(line => 
                    new TextLine { 
                        new TextRun
                        {
                            Style = System.Drawing.FontStyle.Regular,
                            Text = line
                        }});
        }

        FormattedText GetFunctionIntelliSense(IntelliSenseFunctionInfo functionInfo, int currentArgIndex)
        {
            var nameLine = new TextLine { new TextRun { Text = functionInfo.FunctionName + "(" } };
            if (functionInfo.ArgumentList.Count > 0)
            {
                var argNames = functionInfo.ArgumentList.Take(currentArgIndex).Select(arg => arg.ArgumentName).ToArray();
                if (argNames.Length >= 1)
                {
                    nameLine.Add(new TextRun { Text = string.Join(", ", argNames) });
                }

                if (functionInfo.ArgumentList.Count > currentArgIndex)
                {
                    if (argNames.Length >= 1)
                    {
                        nameLine.Add(new TextRun
                        {
                            Text = ", "
                        });
                    }

                    nameLine.Add(new TextRun
                    {
                        Text = functionInfo.ArgumentList[currentArgIndex].ArgumentName,
                        Style = System.Drawing.FontStyle.Bold
                    });

                    argNames = functionInfo.ArgumentList.Skip(currentArgIndex + 1).Select(arg => arg.ArgumentName).ToArray();
                    if (argNames.Length >= 1)
                    {
                        nameLine.Add(new TextRun {Text = ", " + string.Join(", ", argNames)});
                    }
                }
            }
            nameLine.Add(new TextRun { Text = ")" });

            var descriptionLines = GetFunctionDescription(functionInfo);
            
            var formattedText = new FormattedText { nameLine, descriptionLines };
            if (functionInfo.ArgumentList.Count > currentArgIndex)
            {
                formattedText.Add(GetArgumentDescription(functionInfo.ArgumentList[currentArgIndex]));
            }

            return formattedText;
        }

        TextLine GetArgumentDescription(IntelliSenseFunctionInfo.ArgumentInfo argumentInfo)
        {
            return new TextLine { 
                    new TextRun
                    {
                        Style = System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic,
                        Text = argumentInfo.ArgumentName + ": "
                    },
                    new TextRun
                    {
                        Style = System.Drawing.FontStyle.Italic,
                        Text = argumentInfo.Description
                    },
                };
        }

        //public void Shutdown()
        //{
        //    Debug.Print("Shutdown!");
        //    if (_current != null)
        //    {
        //        try
        //        {
        //            _current.Dispose();
        //        }
        //        catch (Exception ex)
        //        {
        //            Debug.Print("!!! Error during Shutdown: " + ex);
        //        }
                
        //        _current = null;
        //    }
        //}

        public void Dispose()
        {
            _syncContextMain.Send(delegate
                {
                    if (_descriptionToolTip != null)
                    {
                        _descriptionToolTip.Dispose();
                        _descriptionToolTip = null;
                    }
                    if (_argumentsToolTip != null)
                    {
                        _argumentsToolTip.Dispose();
                        _argumentsToolTip = null;
                    }
                }, null);

            _syncContextMain = null;
        }

        public void RegisterFunctionInfo(IntelliSenseFunctionInfo functionInfo)
        {
            // TODO : Dictionary from KeyLookup
            _functionInfoMap.Add(functionInfo.FunctionName, functionInfo);
        }
        // TODO: How to UnregisterFunctionInfo ...?
    }
}
