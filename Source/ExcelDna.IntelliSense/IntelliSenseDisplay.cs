using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;

namespace ExcelDna.IntelliSense
{
    class IntelliSenseFunctionInfo
    {
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
    // NOTE: TrackFocus example shows how to do a window 'natively'.
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
        IntPtr _formulaEditWindow;
        IntPtr _functionListWindow;
        
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

        // TODO: Still not sure how to delete / unregister...
        internal void UpdateFunctionInfos(IEnumerable<IntelliSenseFunctionInfo> functionInfos)
        {
            foreach (var fi in functionInfos)
            {
                RegisterFunctionInfo(fi);
            }
        }

        #region Update Preview Filtering
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

            // TODO: Why do this twice....?
            string functionName;
            int currentArgIndex;
            if (FormulaParser.TryGetFormulaInfo(formulaPrefix, out functionName, out currentArgIndex))
            {
                if (_functionInfoMap.ContainsKey(functionName))
                    return true;
            }
            // Not interested...
            Debug.Print($"Not processing formula {formulaPrefix}");
            return false;
        }
        #endregion

        // This runs on the main thread
        void StateUpdate(object sender, UIStateUpdate update)
        {
            Debug.Print($"STATE UPDATE ({update.Update}): \r\t\t\t{update.OldState} \r\t\t=>\t{update.NewState}");

            switch (update.Update)
            {
                case UIStateUpdate.UpdateType.FormulaEditStart:
                    UpdateFormulaEditWindow((update.NewState as UIState.FormulaEdit).FormulaEditWindow);
                    FormulaEditStart();
                    break;
                case UIStateUpdate.UpdateType.FormulaEditMove:
                    var fem = (UIState.FormulaEdit)update.NewState;
                    FormulaEditMove(fem.EditWindowBounds, fem.ExcelTooltipBounds);
                    break;
                case UIStateUpdate.UpdateType.FormulaEditWindowChange:
                    var fewc = (UIState.FormulaEdit)update.NewState;
                    UpdateFormulaEditWindow(fewc.FormulaEditWindow);
                    FormulaEditTextChange(fewc.FormulaPrefix, fewc.EditWindowBounds, fewc.ExcelTooltipBounds);
                    break;
                case UIStateUpdate.UpdateType.FormulaEditTextChange:
                    var fetc = (UIState.FormulaEdit)update.NewState;
                    FormulaEditTextChange(fetc.FormulaPrefix, fetc.EditWindowBounds, fetc.ExcelTooltipBounds);
                    break;
                case UIStateUpdate.UpdateType.FunctionListShow:
                    var fls = (UIState.FunctionList)update.NewState;
                    // TODO: TEMP
                    _functionListWindow = fls.FunctionListWindow;
                    FunctionListShow();
                    FunctionListSelectedItemChange(fls.SelectedItemText, fls.SelectedItemBounds);
                    break;
                case UIStateUpdate.UpdateType.FunctionListMove:
                    var flm = (UIState.FunctionList)update.NewState;
                    FunctionListMove(flm.SelectedItemBounds);
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
                case UIStateUpdate.UpdateType.SelectDataSourceWindowChange:
                case UIStateUpdate.UpdateType.SelectDataSourceHide:
                    // We ignore these
                    break;
                default:
                    throw new InvalidOperationException("Unexpected UIStateUpdate");
            }
        }

        // Runs on the main thread
        void UpdateFormulaEditWindow(IntPtr formulaEditWindow)
        {
            if (_formulaEditWindow != formulaEditWindow)
            {
                _formulaEditWindow = formulaEditWindow;
                if (_argumentsToolTip != null)
                {
                    // Rather ChangeParent...?
                    _argumentsToolTip.Dispose();
                    _argumentsToolTip = null;
                }
                if (_formulaEditWindow != IntPtr.Zero)
                {
                    _argumentsToolTip = new ToolTipForm(_formulaEditWindow);
                    //_argumentsToolTip.OwnerHandle = _formulaEditWindow;
                }
                else
                {
                    // Debug.Fail("Unexpected null FormulaEditWindow...");
                }
            }
        }

        void UpdateFunctionListWindow(IntPtr functionListWindow)
        {
            if (_functionListWindow != functionListWindow)
            {
                _functionListWindow = functionListWindow;
                if (_descriptionToolTip != null)
                {
                    _descriptionToolTip.Dispose();
                    _descriptionToolTip = null;
                }
                if (_functionListWindow != IntPtr.Zero)
                {
                    _descriptionToolTip = new ToolTipForm(_functionListWindow);
                    //_descriptionToolTip.OwnerHandle = _functionListWindow;
                }
                
            }
        }

        // Runs on the main thread
        void FormulaEditStart()
        {
            Debug.Print($"IntelliSenseDisplay - FormulaEditStart");
            if (_formulaEditWindow != IntPtr.Zero && _argumentsToolTip == null)
                _argumentsToolTip = new ToolTipForm(_formulaEditWindow);
        }

        // Runs on the main thread
        void FormulaEditEnd()
        {
            Debug.Print($"IntelliSenseDisplay - FormulaEditEnd");
            // TODO: When can it be null
            if (_argumentsToolTip != null)
            {
                //_argumentsToolTip.Hide();
                _argumentsToolTip.Dispose();
                _argumentsToolTip = null;
            }
        }

        // Runs on the main thread
        void FormulaEditMove(Rect editWindowBounds, Rect excelTooltipBounds)
        {
            Debug.Print($"IntelliSenseDisplay - FormulaEditMove");
            Debug.Assert(_argumentsToolTip != null);
            if (_argumentsToolTip == null)
            {
                Logger.Display.Warn("FormulaEditMode Unexpected null Arguments ToolTip!?");
                return;
            }
            _argumentsToolTip.MoveToolTip((int)editWindowBounds.Left, (int)editWindowBounds.Bottom + 5);
        }

        // Runs on the main thread
        void FormulaEditTextChange(string formulaPrefix, Rect editWindowBounds, Rect excelTooltipBounds)
        {
            Debug.Print($"^^^ FormulaEditStateChanged. CurrentPrefix: {formulaPrefix}, Thread {Thread.CurrentThread.ManagedThreadId}");
            string functionName;
            int currentArgIndex;
            if (FormulaParser.TryGetFormulaInfo(formulaPrefix, out functionName, out currentArgIndex))
            {
                IntelliSenseFunctionInfo functionInfo;
                if (_functionInfoMap.TryGetValue(functionName, out functionInfo))
                {
                    if (_argumentsToolTip != null)
                    {
                        _argumentsToolTip.ShowToolTip(
                            GetFunctionIntelliSense(functionInfo, currentArgIndex),
                            (int)editWindowBounds.Left, (int)editWindowBounds.Bottom + 5);
                    }
                    else
                    {
                        Logger.Display.Warn("FormulaEditTextChange with no arguments tooltip !?");
                    }
                    return;
                }

            }

            // All other paths, we hide the box
            if (_argumentsToolTip != null)
            {
                _argumentsToolTip.Hide();
                //_argumentsToolTip.Dispose();
                //_argumentsToolTip = null;
            }
        }


        // Runs on the main thread
        void FunctionListShow()
        {
            Debug.Print($"IntelliSenseDisplay - FunctionListShow");
            if (_descriptionToolTip == null)
                _descriptionToolTip = new ToolTipForm(_functionListWindow);
        }

        // Runs on the main thread
        void FunctionListHide()
        {
            Debug.Print($"IntelliSenseDisplay - FunctionListHide");
            _descriptionToolTip.Hide();
            //_descriptionToolTip.Dispose();
            //_descriptionToolTip = null;
        }

        // Runs on the main thread
        void FunctionListSelectedItemChange(string selectedItemText, Rect selectedItemBounds)
        {
            Debug.Print($"IntelliSenseDisplay - PopupListSelectedItemChanged - New text - {selectedItemText}, Thread {Thread.CurrentThread.ManagedThreadId}");

            IntelliSenseFunctionInfo functionInfo;
            if (_functionInfoMap.TryGetValue(selectedItemText, out functionInfo))
            {
                // It's ours!
                var descriptionLines = GetFunctionDescriptionOrNull(functionInfo);
                if (descriptionLines != null)
                {
                    _descriptionToolTip.ShowToolTip(
                        text: new FormattedText { GetFunctionDescriptionOrNull(functionInfo) },
                        left: (int)selectedItemBounds.Right + 25,
                        top: (int)selectedItemBounds.Top);
                    return;
                }
            }

            // Not ours or no description
            _descriptionToolTip.Hide();
        }
        
        void FunctionListMove(Rect selectedItemBounds)
        {
            _descriptionToolTip.MoveToolTip((int)selectedItemBounds.Right + 25, (int)selectedItemBounds.Top);
        }

        // TODO: Performance / efficiency - cache these somehow
        // TODO: Probably not a good place for LINQ !?
        static readonly string[] s_newLineStringArray = new string[] { Environment.NewLine };
        IEnumerable<TextLine> GetFunctionDescriptionOrNull(IntelliSenseFunctionInfo functionInfo)
        {
            var description = functionInfo.Description;
            if (string.IsNullOrEmpty(description))
                return null;

            return description.Split(s_newLineStringArray, StringSplitOptions.None)
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

            var descriptionLines = GetFunctionDescriptionOrNull(functionInfo);
            
            var formattedText = new FormattedText { nameLine, descriptionLines };
            if (functionInfo.ArgumentList.Count > currentArgIndex)
            {
                var description = GetArgumentDescription(functionInfo.ArgumentList[currentArgIndex]);
                formattedText.Add(description);
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
                        Text = argumentInfo.Description ?? ""
                    },
                };
        }

        public void Dispose()
        {
            _uiMonitor.StateUpdatePreview -= StateUpdatePreview;
            _uiMonitor.StateUpdate -= StateUpdate;
            //// TODO: How to interact with the pending event callbacks?
            //_syncContextMain.Send(delegate
            //    {
            //        if (_descriptionToolTip != null)
            //        {
            //            _descriptionToolTip.Dispose();
            //            _descriptionToolTip = null;
            //        }
            //        if (_argumentsToolTip != null)
            //        {
            //            _argumentsToolTip.Dispose();
            //            _argumentsToolTip = null;
            //        }
            //    }, null);
            _syncContextMain = null;
        }

        // TODO: Think about case again
        // TODO: Consider locking...
        public void RegisterFunctionInfo(IntelliSenseFunctionInfo functionInfo)
        {
            // TODO : Dictionary from KeyLookup
            IntelliSenseFunctionInfo oldFunctionInfo;
            if (!_functionInfoMap.TryGetValue(functionInfo.FunctionName, out oldFunctionInfo))
            {
                _functionInfoMap.Add(functionInfo.FunctionName, functionInfo);
            }
            else
            {
                // Update  against the function name
                _functionInfoMap[functionInfo.FunctionName] = functionInfo;
            }
        }

        public void UnregisterFunctionInfo(IntelliSenseFunctionInfo functionInfo)
        {
            _functionInfoMap.Remove(functionInfo.FunctionName);
        }
    }
}
