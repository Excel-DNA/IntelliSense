using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
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
        // Not sure we need this here... (the UIMonitor internally uses it, and raises all events on the main thread).
        SynchronizationContext _syncContextMain;
        readonly UIMonitor _uiMonitor;

        readonly Dictionary<string, IntelliSenseFunctionInfo> _functionInfoMap =
            new Dictionary<string, IntelliSenseFunctionInfo>(StringComparer.CurrentCultureIgnoreCase);

        // Need to make these late ...?
        ToolTipForm _descriptionToolTip;
        ToolTipForm _argumentsToolTip;
        
        public IntelliSenseDisplay(SynchronizationContext syncContextMain, UIMonitor uiMonitor)
        {
            // We expect this to be running in a macro context on the main Excel thread (ManagedThreadId = 1).
            #pragma warning disable CS0618 // Type or member is obsolete (GetCurrentThreadId) - But for debugging we want to monitor this anyway
            Debug.Print($"### Thread creating IntelliSenseDisplay: Managed {Thread.CurrentThread.ManagedThreadId}, Native {AppDomain.GetCurrentThreadId()}");
            #pragma warning restore CS0618 // Type or member is obsolete

            _syncContextMain = syncContextMain;
            _uiMonitor = uiMonitor;
            _uiMonitor.StateChanged += _uiMonitor_StateChanged;
        }

        private void _uiMonitor_StateChanged(object sender, UIStateUpdate e)
        {
            Debug.Print($"STATE UPDATE ({e.Update}): {e.OldState} => {e.NewState}");
        }


        //// Runs on the main thread
        //void MainWindowChanged(object _unused_)
        //{
        //    // TODO: This is to guard against shutdown, but we should not have a race here 
        //    //       - shutdown should be on the main thread, as is this event handler.
        //    if (_windowWatcher == null)
        //        return;

        //    // TODO: !!! Reset / re-parent ToolTipWindows
        //    Debug.Print($"IntelliSenseDisplay - MainWindowChanged - New window - {_windowWatcher.MainWindow:X}, Thread {Thread.CurrentThread.ManagedThreadId}");

        //    // _descriptionToolTip.SetOwner(e.Handle); // Not Parent, of course!
        //    if (_descriptionToolTip != null)
        //    {
        //        if (_descriptionToolTip.OwnerHandle != _windowWatcher.MainWindow)
        //        {
        //            _descriptionToolTip.Dispose();
        //            _descriptionToolTip = null;
        //        }
        //    }
        //    if (_argumentsToolTip != null)
        //    {
        //        if (_argumentsToolTip.OwnerHandle != _windowWatcher.MainWindow)
        //        {
        //            _argumentsToolTip.Dispose();
        //            _argumentsToolTip = null;
        //        }
        //    }
        //    // _descriptionToolTip = new ToolTipWindow("", _windowWatcher.MainWindow);
        //}

        //// Runs on the main thread
        //void PopupListSelectedItemChanged(object _unused_)
        //{
        //    Debug.Print($"IntelliSenseDisplay - PopupListSelectedItemChanged - New text - {_popupListWatcher?.SelectedItemText}, Thread {Thread.CurrentThread.ManagedThreadId}");

        //    if (_popupListWatcher == null)
        //        return;
        //    string functionName = _popupListWatcher.SelectedItemText;

        //    IntelliSenseFunctionInfo functionInfo;
        //    if ( _functionInfoMap.TryGetValue(functionName, out functionInfo))
        //    {
        //        if (_descriptionToolTip == null)
        //        {
        //            _descriptionToolTip = new ToolTipForm(_windowWatcher.MainWindow);
        //            _argumentsToolTip = new ToolTipForm(_windowWatcher.MainWindow);
        //        }
        //        // It's ours!
        //        _descriptionToolTip.ShowToolTip(
        //            text: new FormattedText { GetFunctionDescription(functionInfo) }, 
        //            left: (int)_popupListWatcher.SelectedItemBounds.Right + 25,
        //            top:  (int)_popupListWatcher.SelectedItemBounds.Top);
        //    }
        //    else
        //    {
        //        if (_descriptionToolTip != null)
        //        {
        //            _descriptionToolTip.Hide();
        //        }
        //    }
        //}

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

        // TODO: Probably not a good place for LINQ !?
        IEnumerable<TextLine> GetFunctionDescription(IntelliSenseFunctionInfo functionInfo)
        {
            return 
                functionInfo.Description
                .Split(new string[] { Environment.NewLine }, StringSplitOptions.None)
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
