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
    }

    // CONSIDER: Revisit UI Automation Threading: http://msdn.microsoft.com/en-us/library/windows/desktop/ee671692(v=vs.85).aspx
    //           And this threading sample using tlbimp version of Windows 7 native UIA: http://code.msdn.microsoft.com/Windows-7-UI-Automation-6390614a/sourcecode?fileId=21469&pathId=715901329
    class IntelliSenseDisplay : MarshalByRefObject, IDisposable
    {
        static IntelliSenseDisplay _current;
        SynchronizationContext _syncContextMain;
        SynchronizationContext _syncContextAuto;    // On Automation thread.
        
        // NOTE: Add for separate UI Automation Thread
        Thread _threadAuto;
        
        readonly Dictionary<string, IntelliSenseFunctionInfo> _functionInfoMap;

        WindowWatcher _windowWatcher;
        FormulaEditWatcher _formulaEditWatcher;
        PopupListWatcher _popupListWatcher;

        // Need to make these late ...?
        ToolTipForm _descriptionToolTip;
        ToolTipForm _argumentsToolTip;
        
        public IntelliSenseDisplay()
        {
            Debug.Print("### Thread creating IntelliSenseDisplay: " + Thread.CurrentThread.ManagedThreadId);

            _current = this;
            _functionInfoMap = new Dictionary<string, IntelliSenseFunctionInfo>();
            // TODO: Need a separate thread for UI Automation Client - event subscriptions should not be on main UI thread.

            _syncContextMain = new WindowsFormsSynchronizationContext();
            
            // NOTE: Add for separate UI Automation Thread
            _threadAuto = new Thread(RunUIAutomation);
            _threadAuto.Start();

            // NOTE: For running UI Automation Thread on main Excel thread
            //RunUIAutomation();
        }

        public void RegisterFunctionInfo(IntelliSenseFunctionInfo functionInfo)
        {
            // TODO : Dictionary from KeyLookup
            _functionInfoMap.Add(functionInfo.FunctionName, functionInfo);
        }

        void RunUIAutomation()
        {
            // NOTE: Add for separate UI Automation Thread
            _syncContextAuto = new WindowsFormsSynchronizationContext();
            //_syncContextAuto = _syncContextMain;

            _windowWatcher = new WindowWatcher();
            _formulaEditWatcher = new FormulaEditWatcher(_windowWatcher);
            _popupListWatcher = new PopupListWatcher(_windowWatcher);

            _windowWatcher.MainWindowChanged +=
                delegate
                {
                    Debug.Print("### Thread calling MainWindowChanged event: " + Thread.CurrentThread.ManagedThreadId);
                    _syncContextMain.Post(delegate { MainWindowChanged(); }, null);
                    // MainWindowChanged();
                };

            _popupListWatcher.SelectedItemChanged +=
                delegate
                {
                    _syncContextMain.Post(delegate { PopupListSelectedItemChanged(); }, null);
                    // PopupListSelectedItemChanged();
                };
            _formulaEditWatcher.StateChanged +=
                delegate
                {
                    _syncContextMain.Post(delegate { FormulaEditStateChanged(); }, null);
                    // FormulaEditStateChanged();
            
                };

            _windowWatcher.TryInitialize();
            // NOTE: Add for separate UI Automation Thread
             Application.Run();
        }

        void MainWindowChanged()
        {
            // TODO: This is to guard against shutdown, but we should not have a race here 
            //       - shutdown should be on the main thread, as is this event handler.
            if (_windowWatcher == null) return;


            // TODO: !!! Reset / re-parent ToolTipWindows
            Debug.Print("MainWindow Change - " + _windowWatcher.MainWindow.ToString("X"));
            Debug.Print("### Thread calling MainWindowChanged method: " + Thread.CurrentThread.ManagedThreadId);

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

        void PopupListSelectedItemChanged()
        {
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
                    new FormattedText { GetFunctionDescription(functionInfo) }, 
                    (int)_popupListWatcher.SelectedItemBounds.Right + 25, (int)_popupListWatcher.SelectedItemBounds.Top);
            }
            else
            {
                if (_descriptionToolTip != null)
                {
                    _descriptionToolTip.Hide();
                }
            }
        }

        // TODO: Need better formula parsing story here
        // Here are some ideas: http://fastexcel.wordpress.com/2013/10/27/parsing-functions-from-excel-formulas-using-vba-is-mid-or-a-byte-array-the-best-method/
        void FormulaEditStateChanged()
        {
            // Check for watcher already disposed 
            // CONSIDER: How to manage threading with disposal...?
            if (_formulaEditWatcher == null) return;
            Debug.Print("^^^ FormulaEditStateChanged. CurrentPrefix: " + _formulaEditWatcher.CurrentPrefix);
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

        public static void Shutdown()
        {
            Debug.Print("Shutdown!");
            if (_current != null)
            {
                try
                {
                    _current.Dispose();
                }
                catch (Exception ex)
                {
                    Debug.Print("!!! Error during Shutdown: " + ex);
                }
                
                _current = null;
            }
        }

        public void Dispose()
        {
            _current._syncContextAuto.Send(delegate
                {
                    if (_windowWatcher != null)
                    {
                        _windowWatcher.Dispose();
                        _windowWatcher = null;
                    }
                    if (_formulaEditWatcher != null)
                    {
                        _formulaEditWatcher.Dispose();
                        _formulaEditWatcher = null;
                    }
                    if (_popupListWatcher != null)
                    {
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

            // NOTE: Add for separate UI Automation Thread
            _threadAuto.Abort();
            _threadAuto = null;
            _syncContextAuto = null;
            _syncContextMain = null;
        }
    }
}
