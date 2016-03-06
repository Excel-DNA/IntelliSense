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
    class IntelliSenseDisplay : MarshalByRefObject, IDisposable
    {
        SynchronizationContext _syncContextMain;                // Running on the main Excel thread (not a 'macro' context, though)
        SingleThreadSynchronizationContext _syncContextAuto;    // Running on the Automation thread.

        readonly Dictionary<string, IntelliSenseFunctionInfo> _functionInfoMap =
            new Dictionary<string, IntelliSenseFunctionInfo>(StringComparer.CurrentCultureIgnoreCase);

        WindowWatcher _windowWatcher;
        FormulaEditWatcher _formulaEditWatcher;
        PopupListWatcher _popupListWatcher;
        //SelectDataSourceWatcher _selectDataSourceWatcher;

        // Need to make these late ...?
        ToolTipForm _descriptionToolTip;
        ToolTipForm _argumentsToolTip;
        
        public IntelliSenseDisplay()
        {
            // We expect this to be running in a macro context on the main Excel thread (ManagedThreadId = 1).
            Debug.Print($"### Thread creating IntelliSenseDisplay: Managed {Thread.CurrentThread.ManagedThreadId}, Native {AppDomain.GetCurrentThreadId()}");

            _syncContextMain = new WindowsFormsSynchronizationContext();

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

            _windowWatcher = new WindowWatcher(_syncContextAuto);
            _formulaEditWatcher = new FormulaEditWatcher(_windowWatcher, _syncContextAuto);
            _popupListWatcher = new PopupListWatcher(_windowWatcher, _syncContextAuto);
            // _selectDataSourceWatcher = new SelectDataSourceWatcher(_windowWatcher, _syncContextAuto);

            _windowWatcher.MainWindowChanged += _windowWatcher_MainWindowChanged;
            _popupListWatcher.SelectedItemChanged += _popupListWatcher_SelectedItemChanged;
            _formulaEditWatcher.StateChanged += _formulaEditWatcher_StateChanged;

            _windowWatcher.TryInitialize();

            _syncContextAuto.RunOnCurrentThread();
        }

        // Runs on the auto thread
        void _windowWatcher_MainWindowChanged(object sender, EventArgs args)
        {
            _syncContextMain.Post(MainWindowChanged, null);
        }

        // Runs on the auto thread
        void _popupListWatcher_SelectedItemChanged(object sender, EventArgs args)
        {
            _syncContextMain.Post(PopupListSelectedItemChanged, null);
        }

        // Runs on the auto thread
        void _formulaEditWatcher_StateChanged(object sender, FormulaEditWatcher.StateChangeEventArgs args)
        {
            _syncContextMain.Post(FormulaEditStateChanged, args.StateChangeType);
        }

        // Runs on the main thread
        void MainWindowChanged(object _unused_)
        {
            // TODO: This is to guard against shutdown, but we should not have a race here 
            //       - shutdown should be on the main thread, as is this event handler.
            if (_windowWatcher == null) return;

            // TODO: !!! Reset / re-parent ToolTipWindows
            Debug.Print($"IntelliSenseDisplay - MainWindowChanged - New window - {_windowWatcher.MainWindow:X}, Thread {Thread.CurrentThread.ManagedThreadId}");

            // _descriptionToolTip.SetOwner(e.Handle); // Not Parent, of course!
            if (_descriptionToolTip != null)
            {
                if (_descriptionToolTip.OwnerHandle != _windowWatcher.MainWindow)
                {
                    _descriptionToolTip.Dispose();
                    _descriptionToolTip = null;
                }
            }
            if (_argumentsToolTip != null)
            {
                if (_argumentsToolTip.OwnerHandle != _windowWatcher.MainWindow)
                {
                    _argumentsToolTip.Dispose();
                    _argumentsToolTip = null;
                }
            }
            // _descriptionToolTip = new ToolTipWindow("", _windowWatcher.MainWindow);
        }

        // Runs on the main thread
        void PopupListSelectedItemChanged(object _unused_)
        {
            Debug.Print($"IntelliSenseDisplay - PopupListSelectedItemChanged - New text - {_popupListWatcher?.SelectedItemText}, Thread {Thread.CurrentThread.ManagedThreadId}");

            if (_popupListWatcher == null) return;
            string functionName = _popupListWatcher.SelectedItemText;

            IntelliSenseFunctionInfo functionInfo;
            if ( _functionInfoMap.TryGetValue(functionName, out functionInfo))
            {
                if (_descriptionToolTip == null)
                {
                    _descriptionToolTip = new ToolTipForm(_windowWatcher.MainWindow);
                    _argumentsToolTip = new ToolTipForm(_windowWatcher.MainWindow);
                }
                // It's ours!
                _descriptionToolTip.ShowToolTip(
                    text: new FormattedText { GetFunctionDescription(functionInfo) }, 
                    left: (int)_popupListWatcher.SelectedItemBounds.Right + 25,
                    top:  (int)_popupListWatcher.SelectedItemBounds.Top);
            }
            else
            {
                if (_descriptionToolTip != null)
                {
                    _descriptionToolTip.Hide();
                }
            }
        }

        // Runs on the main thread
        // TODO: Need better formula parsing story here
        // Here are some ideas: http://fastexcel.wordpress.com/2013/10/27/parsing-functions-from-excel-formulas-using-vba-is-mid-or-a-byte-array-the-best-method/
        void FormulaEditStateChanged(object stateChangeTypeObj)
        {
            var stateChangeType = (FormulaEditWatcher.StateChangeType)stateChangeTypeObj;
            // Check for watcher already disposed 
            // CONSIDER: How to manage threading with disposal...?
            if (_formulaEditWatcher == null) return;

            if (stateChangeType == FormulaEditWatcher.StateChangeType.Move && _argumentsToolTip != null)
            {
                _argumentsToolTip.MoveToolTip(
                    (int)_formulaEditWatcher.EditWindowBounds.Left, (int)_formulaEditWatcher.EditWindowBounds.Bottom + 5);
                return;
            }

            Debug.Print($"^^^ FormulaEditStateChanged. CurrentPrefix: {_formulaEditWatcher.CurrentPrefix}, Thread {Thread.CurrentThread.ManagedThreadId}");
            if (_formulaEditWatcher.IsEditingFormula && _formulaEditWatcher.CurrentPrefix != null)
            {
                string prefix = _formulaEditWatcher.CurrentPrefix;
                var match = Regex.Match(prefix, @"^=(?<functionName>\w*)\(");
                if (match.Success)
                {
                    string functionName = match.Groups["functionName"].Value;

                    IntelliSenseFunctionInfo functionInfo;
                    if (_functionInfoMap.TryGetValue(functionName, out functionInfo))
                    {
                        // It's ours!
                        if (_argumentsToolTip == null)
                        {
                            _argumentsToolTip = new ToolTipForm(_windowWatcher.MainWindow);
                        }

                        // TODO: Fix this: Need to consider subformulae
                        int currentArgIndex = _formulaEditWatcher.CurrentPrefix.Count(c => c == ',');
                        _argumentsToolTip.ShowToolTip(
                            GetFunctionIntelliSense(functionInfo, currentArgIndex),
                            (int)_formulaEditWatcher.EditWindowBounds.Left, (int)_formulaEditWatcher.EditWindowBounds.Bottom + 5);
                        return;
                    }
                }
            }

            // All other paths, we just clear the box
            if (_argumentsToolTip != null)
                _argumentsToolTip.Hide();
        }

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
            if (_syncContextAuto == null)
                return;

            _syncContextAuto.Send(delegate
                {
                    if (_windowWatcher != null)
                    {
                        _windowWatcher.MainWindowChanged -= _windowWatcher_MainWindowChanged;
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

            _syncContextAuto.Complete();
            // CONSIDER: Maybe wait for the _syncContextAuto to finish...?
            _syncContextAuto = null;

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
