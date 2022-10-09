using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Windows;
using XlCallInt = ExcelDna.Integration.XlCall;

namespace ExcelDna.IntelliSense
{
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

        readonly Dictionary<string, FunctionInfo> _functionInfoMap =
            new Dictionary<string, FunctionInfo>(StringComparer.CurrentCultureIgnoreCase);

        // Need to make these late ...?
        ToolTipForm _descriptionToolTip;
        ToolTipForm _argumentsToolTip;
        IntPtr _formulaEditWindow;
        IntPtr _functionListWindow;
        string _argumentSeparator = ", ";
        
        const int DescriptionLeftMargin = 3;

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

            InitializeOptions();
        }
        
        // Runs on the main Excel thread in a macro context.
        void InitializeOptions()
        {
            Logger.Display.Verbose("InitializeOptions");
            string listSeparator = ",";
            string standardFontName = "Calibri";
            double standardFontSize = 11.0;
            object result;
            if (XlCallInt.TryExcel(XlCallInt.xlfGetWorkspace, out result, 37) == XlCallInt.XlReturn.XlReturnSuccess)
            {
                object[,] options = result as object[,];
                if (options != null)
                {
                    listSeparator = (string)options[0, 4];
                    Logger.Initialization.Verbose($"InitializeOptions - Set ListSeparator to {listSeparator}");
                }
            }
            FormulaParser.ListSeparator = listSeparator[0];
            _argumentSeparator = listSeparator + " ";

            if (XlCallInt.TryExcel(XlCallInt.xlfGetWorkspace, out result, 56) == XlCallInt.XlReturn.XlReturnSuccess)
            {
                // Standard Font Name
                standardFontName = (string)result;
                Logger.Initialization.Verbose($"InitializeOptions - Set StandardFontName to {standardFontName}");
            }
            if (XlCallInt.TryExcel(XlCallInt.xlfGetWorkspace, out result, 57) == XlCallInt.XlReturn.XlReturnSuccess)
            {
                // Standard Font Size
                standardFontSize = (double)result;
                Logger.Initialization.Verbose($"InitializeOptions - Set StandardFontSize to {standardFontSize}");
            }
            ToolTipForm.SetStandardFont(standardFontName, standardFontSize);
        }

        // TODO: Still not sure how to delete / unregister...
        internal void UpdateFunctionInfos(IEnumerable<FunctionInfo> functionInfos)
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
            else if (update.Update == UIStateUpdate.UpdateType.FormulaEditExcelToolTipChange)
            {
                // If if has just been hidden, we need no extra processing and can skip the Main thread call
                var fe = (UIState.FormulaEdit)update.NewState;
                enable = fe.ExcelToolTipWindow != IntPtr.Zero;
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
                    var fes = (UIState.FormulaEdit)update.NewState;
                    UpdateFormulaEditWindow(fes.FormulaEditWindow);
                    FormulaEditStart(fes.FormulaPrefix, fes.EditWindowBounds, fes.ExcelToolTipWindow);
                    break;
                case UIStateUpdate.UpdateType.FormulaEditMove:
                    var fem = (UIState.FormulaEdit)update.NewState;
                    FormulaEditMove(fem.EditWindowBounds, fem.ExcelToolTipWindow);
                    break;
                case UIStateUpdate.UpdateType.FormulaEditWindowChange:
                    var fewc = (UIState.FormulaEdit)update.NewState;
                    UpdateFormulaEditWindow(fewc.FormulaEditWindow);
                    FormulaEditTextChange(fewc.FormulaPrefix, fewc.EditWindowBounds, fewc.ExcelToolTipWindow);
                    break;
                case UIStateUpdate.UpdateType.FormulaEditTextChange:
                    var fetc = (UIState.FormulaEdit)update.NewState;
                    FormulaEditTextChange(fetc.FormulaPrefix, fetc.EditWindowBounds, fetc.ExcelToolTipWindow);
                    break;
                case UIStateUpdate.UpdateType.FormulaEditExcelToolTipChange:
                    var fett = (UIState.FormulaEdit)update.NewState;
                    FormulaEditExcelToolTipShow(fett.EditWindowBounds, fett.ExcelToolTipWindow);
                    break;
                case UIStateUpdate.UpdateType.FunctionListShow:
                    var fls = (UIState.FunctionList)update.NewState;
                    _functionListWindow = fls.FunctionListWindow;
                    FunctionListShow();
                    FunctionListSelectedItemChange(fls.SelectedItemText, fls.SelectedItemBounds, fls.FunctionListBounds);
                    break;
                case UIStateUpdate.UpdateType.FunctionListMove:
                    var flm = (UIState.FunctionList)update.NewState;
                    FunctionListMove(flm.SelectedItemBounds, flm.FunctionListBounds);
                    break;
                case UIStateUpdate.UpdateType.FunctionListSelectedItemChange:
                    var fl = (UIState.FunctionList)update.NewState;
                    FunctionListSelectedItemChange(fl.SelectedItemText, fl.SelectedItemBounds, fl.FunctionListBounds);
                    break;
                case UIStateUpdate.UpdateType.FunctionListHide:
                case UIStateUpdate.UpdateType.FunctionListWindowChange:
                    FunctionListHide();
                    break;
                case UIStateUpdate.UpdateType.FormulaEditEnd:
                    FormulaEditEnd();
                    break;
                    
                case UIStateUpdate.UpdateType.SelectDataSourceShow:
                case UIStateUpdate.UpdateType.SelectDataSourceWindowChange:
                case UIStateUpdate.UpdateType.SelectDataSourceHide:
                    // We ignore these for now
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

        //void UpdateFunctionListWindow(IntPtr functionListWindow)
        //{
        //    if (_functionListWindow != functionListWindow)
        //    {
        //        _functionListWindow = functionListWindow;
        //        if (_descriptionToolTip != null)
        //        {
        //            _descriptionToolTip.Dispose();
        //            _descriptionToolTip = null;
        //        }
        //        if (_functionListWindow != IntPtr.Zero)
        //        {
        //            _descriptionToolTip = new ToolTipForm(_functionListWindow);
        //            //_descriptionToolTip.OwnerHandle = _functionListWindow;
        //        }
                
        //    }
        //}

        // Runs on the main thread
        void FormulaEditStart(string formulaPrefix, Rect editWindowBounds, IntPtr excelToolTipWindow)
        {
            Debug.Print($"IntelliSenseDisplay - FormulaEditStart - FormulaEditWindow: {_formulaEditWindow}, ArgumentsToolTip: {_argumentsToolTip}");
            if (_formulaEditWindow != IntPtr.Zero && _argumentsToolTip == null)
                _argumentsToolTip = new ToolTipForm(_formulaEditWindow);

            // Normally we would have no formula at this point.
            // One exception is after mouse-click on the formula list, we then need to process it.
            if (!string.IsNullOrEmpty(formulaPrefix))
                FormulaEditTextChange(formulaPrefix, editWindowBounds, excelToolTipWindow);
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
        void FormulaEditMove(Rect editWindowBounds, IntPtr excelToolTipWindow)
        {
            Debug.Print($"IntelliSenseDisplay - FormulaEditMove");
            if (_argumentsToolTip == null)
            {
                Logger.Display.Warn("FormulaEditMode Unexpected null Arguments ToolTip!?");
                return;
            }
            int topOffset = GetTopOffset(excelToolTipWindow);
            try
            {
                _argumentsToolTip.MoveToolTip((int)editWindowBounds.Left, (int)editWindowBounds.Bottom + 5, topOffset);
            }
            catch (Exception ex)
            {
                Logger.Display.Warn($"IntelliSenseDisplay - FormulaEditMove Error - {ex}");
                // Recycle the Arguments ToolTip - won't show now, but should for the next function
                _argumentsToolTip.Dispose();
                _argumentsToolTip = null;
            }
        }

        // Runs on the main thread
        void FormulaEditTextChange(string formulaPrefix, Rect editWindowBounds, IntPtr excelToolTipWindow)
        {
            Debug.Print($"^^^ FormulaEditStateChanged. CurrentPrefix: {formulaPrefix}, Thread {Thread.CurrentThread.ManagedThreadId}");
            string functionName;
            int currentArgIndex;
            if (FormulaParser.TryGetFormulaInfo(formulaPrefix, out functionName, out currentArgIndex))
            {
                FunctionInfo functionInfo;
                if (_functionInfoMap.TryGetValue(functionName, out functionInfo))
                {
                    var lineBeforeFunctionName = FormulaParser.GetLineBeforeFunctionName(formulaPrefix, functionName);
                    // We have a function name and we want to show info
                    if (_argumentsToolTip != null)
                    {
                        // NOTE: Hiding or moving just once doesn't help - the tooltip pops up in its original place again
                        // TODO: Try to move it off-screen, behind or make invisible
                        //if (!_argumentsToolTip.Visible)
                        //{
                        //    // Fiddle a bit with the ExcelToolTip if it is already visible when we first show our FunctionEdit ToolTip
                        //    // At other times, the explicit UI update should catch and hide as appropriate
                        //    if (excelToolTipWindow != IntPtr.Zero)
                        //    {
                        //        Win32Helper.HideWindow(excelToolTipWindow);
                        //    }
                        //}
                        int topOffset = GetTopOffset(excelToolTipWindow);
                        FormattedText infoText = GetFunctionIntelliSense(functionInfo, currentArgIndex);
                        try
                        {
                            _argumentsToolTip.ShowToolTip(infoText, lineBeforeFunctionName, (int)editWindowBounds.Left, (int)editWindowBounds.Bottom + 5, topOffset);
                        }
                        catch (Exception ex)
                        {
                            Logger.Display.Warn($"IntelliSenseDisplay - FormulaEditTextChange Error - {ex}");
                            _argumentsToolTip.Dispose();
                            _argumentsToolTip = null;
                        }
                    }
                    else
                    {
                        Logger.Display.Info("FormulaEditTextChange with no arguments tooltip !?");
                    }
                    return;
                }

            }

            // All other paths, we hide the box
           _argumentsToolTip?.Hide();
        }


        // This helper just keeps us out of the Excel tooltip's way.
        int GetTopOffset(IntPtr excelToolTipWindow)
        {
            // TODO: Maybe get its height...?
            return (excelToolTipWindow == IntPtr.Zero) ? 0 : 18;
        }

        void FormulaEditExcelToolTipShow(Rect editWindowBounds, IntPtr excelToolTipWindow)
        {
            //    // Excel tool tip has just been shown
            //    // If we're showing the arguments dialog, hide the Excel tool tip
            //    if (_argumentsToolTip != null && _argumentsToolTip.Visible)
            //    {
            //        Win32Helper.HideWindow(excelToolTipWindow);
            //    }

            if (_argumentsToolTip != null && _argumentsToolTip.Visible)
            {
                int topOffset = GetTopOffset(excelToolTipWindow);
                try
                {
                    _argumentsToolTip.MoveToolTip((int)editWindowBounds.Left, (int)editWindowBounds.Bottom + 5, topOffset);
                }
                catch (Exception ex)
                {
                    Logger.Display.Warn($"IntelliSenseDisplay - FormulaEditExcelToolTipShow Error - {ex}");
                    _argumentsToolTip.Dispose();
                    _argumentsToolTip = null;
                }
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
            _descriptionToolTip?.Hide();
        }

        // Runs on the main thread
        void FunctionListSelectedItemChange(string selectedItemText, Rect selectedItemBounds, Rect listBounds)
        {
            Logger.Display.Verbose($"IntelliSenseDisplay - PopupListSelectedItemChanged - {selectedItemText} List/Item Bounds: {listBounds} / {selectedItemBounds}");

            FunctionInfo functionInfo;
            if (!string.IsNullOrEmpty(selectedItemText) &&
                _functionInfoMap.TryGetValue(selectedItemText, out functionInfo))
            {
                // It's ours!
                var descriptionLines = GetFunctionDescriptionOrNull(functionInfo);
                if (descriptionLines != null)
                {
                    try
                    {
                        _descriptionToolTip?.ShowToolTip(
                            text: new FormattedText { descriptionLines },
                            linePrefix: null,
                            left: (int)listBounds.Right + DescriptionLeftMargin,
                            top: (int)selectedItemBounds.Bottom - 18,
                            topOffset: 0,
                            listLeft: (int)selectedItemBounds.Left,
                            listTop: (int)selectedItemBounds.Top);
                        return;
                    }
                    catch (Exception ex)
                    {
                        Logger.Display.Warn($"IntelliSenseDisplay - PopupListSelectedItemChanged Error - {ex}");
                        // Recycle the _DescriptionToolTip - won't show now, but should for the next function
                        _descriptionToolTip.Dispose();
                        _descriptionToolTip = null;
                        return;
                    }
                }
            }

            // Not ours or no description
            _descriptionToolTip?.Hide();
        }
        
        void FunctionListMove(Rect selectedItemBounds, Rect listBounds)
        {
            try
            {
                _descriptionToolTip?.MoveToolTip(
                   left: (int)listBounds.Right + DescriptionLeftMargin,
                   top: (int)selectedItemBounds.Bottom - 18,
                   topOffset: 0,
                   listLeft: (int)selectedItemBounds.Left,
                   listTop: (int)selectedItemBounds.Top);
            }
            catch (Exception ex)
            {
                Logger.Display.Warn($"IntelliSenseDisplay - FunctionListMove Error - {ex}");
                // Recycle the _DescriptionToolTip - won't show now, but should for the next function
                _descriptionToolTip?.Dispose();
                _descriptionToolTip = null;
            }
        }

        // TODO: Performance / efficiency - cache these somehow
        // TODO: Probably not a good place for LINQ !?
        static readonly string[] s_newLineStringArray = new string[] { Environment.NewLine };
        IEnumerable<TextLine> GetFunctionDescriptionOrNull(FunctionInfo functionInfo)
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

        FormattedText GetFunctionIntelliSense(FunctionInfo functionInfo, int currentArgIndex)
        {
            // In case of the special params pattern (x, y, arg1, ...) we base the argument display on an expanded argument list, matching Excel's behaviour,
            // and the magic expansion in the function wizard.
            var argumentList = GetExpandedArgumentList(functionInfo, currentArgIndex);

            var nameLine = new TextLine { new TextRun { Text = functionInfo.Name, LinkAddress = FixHelpTopic(functionInfo.HelpTopic) } };
            nameLine.Add(new TextRun { Text = "(" });
            if (argumentList.Count > 0)
            {
                var argNames = argumentList.Take(currentArgIndex).Select(arg => arg.Name).ToArray();
                if (argNames.Length >= 1)
                {
                    nameLine.Add(new TextRun { Text = string.Join(_argumentSeparator, argNames) });
                }

                if (argumentList.Count > currentArgIndex)
                {
                    if (argNames.Length >= 1)
                    {
                        nameLine.Add(new TextRun
                        {
                            Text = _argumentSeparator
                        });
                    }

                    nameLine.Add(new TextRun
                    {
                        Text = argumentList[currentArgIndex].Name,
                        Style = System.Drawing.FontStyle.Bold
                    });

                    argNames = argumentList.Skip(currentArgIndex + 1).Select(arg => arg.Name).ToArray();
                    if (argNames.Length >= 1)
                    {
                        nameLine.Add(new TextRun {Text = _argumentSeparator + string.Join(_argumentSeparator, argNames)});
                    }
                }
            }
            nameLine.Add(new TextRun { Text = ")" });

            var descriptionLines = GetFunctionDescriptionOrNull(functionInfo);

            var formattedText = new FormattedText { nameLine, descriptionLines };
            if (argumentList.Count > currentArgIndex)
            {
                var description = GetArgumentDescription(argumentList[currentArgIndex]);
                formattedText.Add(description);
            }

            return formattedText;
        }

        // In case of the special params pattern (x, y, arg1, ...) we base the argument display on an expanded argument list, matching Excel's behaviour,
        // and the magic expansion in the function wizard.
        // Thanks to @amibar for figuring this out.
        // NOTE: We might need to get the whole formula, the current location (or prefix) and the currentArgIndex to implement Excel's behaviour for params parameters.
        // Usually just having the prefix is OK, but in case we have the formula: F(params object[] args) and we write in the formula editor =F(1,2,3,4,5,6,7)
        // and then we move the cursor to point on the second argument, our current implementation will shorten its text and omit any argument after the 3rd argument.
        // But Excel will keep showing the vurtual argument list corresponding to the full formula.
        // There is no technical problem in getting the full formula - PenHelper will give us the required info - but tracking this throughout the IntelliSense state 
        // affects the code in a lot of places, and the benefits seem small, particularly in this case of quirky Excel behaviour.
        List<FunctionInfo.ArgumentInfo> GetExpandedArgumentList(FunctionInfo functionInfo, int currentArgIndex)
        {
            // Note: Using params for the last argument
            if (functionInfo.ArgumentList.Count > 1 &&
                functionInfo.ArgumentList[functionInfo.ArgumentList.Count - 1].Name == "..." &&
                functionInfo.ArgumentList[functionInfo.ArgumentList.Count - 2].Name.EndsWith("1")) // Note: Need both the Arg1 and the ... to trigger the expansion in the function wizard?
            {
                var paramsIndex = functionInfo.ArgumentList.Count - 2;

                // Take the last named argument and omit the "1" that the registration added
                var paramsDesc = functionInfo.ArgumentList[paramsIndex].Description;
                string paramsBaseName = functionInfo.ArgumentList[paramsIndex].Name.TrimEnd('1');
                int currentParamsArgIndex = currentArgIndex - paramsIndex;

                // Omit last two entries ("Arg1" and "...") from the original argument list
                var newList = new List<FunctionInfo.ArgumentInfo>(functionInfo.ArgumentList.Take(paramsIndex));
                // Add back with parens the "[Arg1]"
                newList.Add(new FunctionInfo.ArgumentInfo { Name = $"[{paramsBaseName + 1}]", Description = paramsDesc });

                // Now if we're at Arg4 (currentParamsIndex=3) then add "[Arg2],[Arg3],[Arg4],[Arg5],..."
                for (int i = 2; i <= currentParamsArgIndex + 2; i++)
                {
                    newList.Add(new FunctionInfo.ArgumentInfo { Name = $"[{paramsBaseName + i}]", Description = paramsDesc });
                }
                newList.Add(new FunctionInfo.ArgumentInfo { Name = "...", Description = paramsDesc });

                return newList;
            }
            else
            {
                // No problems - just return the real argumentlist
                return functionInfo.ArgumentList;
            }
        }

        IEnumerable<TextLine> GetArgumentDescription(FunctionInfo.ArgumentInfo argumentInfo)
        {
            if (string.IsNullOrEmpty(argumentInfo.Description))
                yield break;

            var lines = argumentInfo.Description.Split(s_newLineStringArray, StringSplitOptions.None);
            yield return new TextLine {
                    new TextRun
                    {
                        Style = System.Drawing.FontStyle.Bold | System.Drawing.FontStyle.Italic,
                        Text = argumentInfo.Name + ": "
                    },
                    new TextRun
                    {
                        Style = System.Drawing.FontStyle.Italic,
                        Text = lines.FirstOrDefault() ?? ""
                    },
                };

            foreach (var line in lines.Skip(1))
            {
                yield return new TextLine {
                    new TextRun
                    {
                        Style = System.Drawing.FontStyle.Italic,
                        Text = line
                    }};
            }
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

        // Removes the !0 that we add to make Excel happy
        string FixHelpTopic(string helpTopic)
        {
            if (helpTopic != null && helpTopic.EndsWith("!0"))
                return helpTopic.Substring(0, helpTopic.Length - 2);
            return helpTopic;
        }

        // TODO: Think about case again
        // TODO: Consider locking...
        public void RegisterFunctionInfo(FunctionInfo functionInfo)
        {
            // TODO : Dictionary from KeyLookup
            FunctionInfo oldFunctionInfo;
            if (!_functionInfoMap.TryGetValue(functionInfo.Name, out oldFunctionInfo))
            {
                _functionInfoMap.Add(functionInfo.Name, functionInfo);
            }
            else
            {
                // Update against the function name
                _functionInfoMap[functionInfo.Name] = functionInfo;
            }
        }

        public void UnregisterFunctionInfo(FunctionInfo functionInfo)
        {
            _functionInfoMap.Remove(functionInfo.Name);
        }
    }
}
