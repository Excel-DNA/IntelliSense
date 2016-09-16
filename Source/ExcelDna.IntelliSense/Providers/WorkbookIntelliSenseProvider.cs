using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna.IntelliSense
{
    // For VBA code, (either in a regular workbook that is open, or in an add-in)
    // we allow the IntelliSense info to be put into a worksheet (possibly hidden or very hidden)
    // In the Workbook that contains the VBA.
    // Initially we won't scope the IntelliSense to the Workbook where the UDFs are defined, 
    // but we should consider that.

    // TODO: Can't we read the Application.MacroOptions...?

    class WorkbookIntelliSenseProvider : IIntelliSenseProvider
    {
        const string intelliSenseWorksheetName = "_IntelliSense_";
        const string functionInfoId = "FunctionInfo";

        class WorkbookRegistrationInfo
        {
            readonly string _name;
            DateTime _lastUpdate;               // Version indicator to enumerate from scratch
            object[,] _regInfo = null;          // Default value
            string _path;

            public WorkbookRegistrationInfo(string name)
            {
                _name = name;
            }
 
            // Called in a macro context
            public void Refresh()
            {
                var app = (Application)ExcelDnaUtil.Application;
                var wb = app.Workbooks[_name];
                _path = wb.Path;

                try
                {
                    var ws = wb.Sheets[intelliSenseWorksheetName];

                    var info = ws.UsedRange.Value;
                    _regInfo = info as object[,];
                }
                catch (Exception ex)
                {
                    // We expect this if there is no sheet.
                    Debug.Print("WorkbookIntelliSenseProvider.Refresh Error : " + ex.Message);
                    _regInfo = null;
                }
                _lastUpdate = DateTime.Now;
            }

            // Not in macro context - don't call Excel, could be any thread.
            public IEnumerable<FunctionInfo> GetFunctionInfos()
            {
                // to avoid worries about locking and this being updated from another thread, we take a copy of the _regInfo reference
                var regInfo = _regInfo;

                if (regInfo == null)
                    yield break;

                int numRows = regInfo.GetLength(0);
                int numCols = regInfo.GetLength(1);
                if (numRows < 1 || numCols < 1)
                    yield break;

                var idCheck = regInfo[1, 1] as string;
                if (!functionInfoId.Equals(idCheck, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Provider.Info($"Workbook - Invalid FunctionInfo Identifier: ({idCheck})");
                    yield break;
                }

                // regInfo is 1-based: object[1..x, 1..y].
                for (int i = 2; i <= numRows; i++)
                {
                    string functionName = regInfo[i, 1] as string;
                    string description = regInfo[i, 2] as string;
                    string helpTopic = regInfo[i, 3] as string;

                    List<FunctionInfo.ArgumentInfo> argumentInfos = new List<FunctionInfo.ArgumentInfo>();
                    for (int j = 4; j <= numCols - 1; j += 2)
                    {
                        var arg = regInfo[i, j] as string;
                        var argDesc = regInfo[i, j + 1] as string;
                        if (!string.IsNullOrEmpty(arg))
                        {
                            argumentInfos.Add(new FunctionInfo.ArgumentInfo
                            {
                                Name = arg,
                                Description = argDesc
                            });
                        }
                    }

                    helpTopic = FunctionInfo.ExpandHelpTopic(_path, helpTopic);

                    yield return new FunctionInfo
                    {
                        Name = functionName,
                        Description = description,
                        HelpTopic = helpTopic,
                        ArgumentList = argumentInfos,
                        SourcePath = _name
                    };
                }
            }
        }

        Dictionary<string, WorkbookRegistrationInfo> _workbookRegistrationInfos = new Dictionary<string, WorkbookRegistrationInfo>();
        XmlIntelliSenseProvider _xmlProvider;
        public event EventHandler Invalidate;

        #region IIntelliSenseProvider implementation

        public WorkbookIntelliSenseProvider()
        {
            _xmlProvider = new XmlIntelliSenseProvider();
            _xmlProvider.Invalidate += ( sender, e) => OnInvalidate();
        }

        public void Initialize()
        {
            Logger.Provider.Info("WorkbookIntelliSenseProvider.Initialize");
            _xmlProvider.Initialize();

            // The events are just to keep track of the set of open workbooks, 
            var xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.WorkbookOpen += Excel_WorkbookOpen;
            xlApp.WorkbookBeforeClose += Excel_WorkbookBeforeClose;
            xlApp.WorkbookAddinInstall += Excel_WorkbookAddinInstall;
            xlApp.WorkbookAddinUninstall += Excel_WorkbookAddinUninstall;

            lock (_workbookRegistrationInfos)
            {
                foreach (Workbook wb in xlApp.Workbooks)
                {
                    var name = wb.Name;
                    if (!_workbookRegistrationInfos.ContainsKey(name))
                    {
                        WorkbookRegistrationInfo regInfo = new WorkbookRegistrationInfo(name);
                        _workbookRegistrationInfos[name] = regInfo;

                        regInfo.Refresh();

                        RegisterWithXmlProvider(wb);
                    }
                }
                if (ExcelDnaUtil.ExcelVersion >= 14.0)
                {
                    foreach (AddIn addIn in xlApp.AddIns2)
                    {
                        if (addIn.IsOpen && Path.GetExtension(addIn.FullName) != ".xll")
                        {
                            // Can it be "Open" and not be loaded?
                            var name = addIn.Name;
                            Workbook wbAddIn;
                            try
                            {
                                // TODO: Log
                                wbAddIn = xlApp.Workbooks[name];
                            }
                            catch
                            {
                                // TODO: Log
                                continue;
                            }

                            WorkbookRegistrationInfo regInfo = new WorkbookRegistrationInfo(name);
                            _workbookRegistrationInfos[name] = regInfo;

                            regInfo.Refresh();

                            RegisterWithXmlProvider(wbAddIn);
                        }
                    }
                }
            }
        }

        // Runs on the main thread
        public void Refresh()
        {
            Logger.Provider.Info("WorkbookIntelliSenseProvider.Refresh");
            lock (_workbookRegistrationInfos)
            {
                foreach (var regInfo in _workbookRegistrationInfos.Values)
                {
                    regInfo.Refresh();
                }
                _xmlProvider.Refresh();
            }
        }

        // May be called from any thread
        public IList<FunctionInfo> GetFunctionInfos()
        {
            lock (_workbookRegistrationInfos)
            {
                var workbookInfos = _workbookRegistrationInfos.Values.SelectMany(ri => ri.GetFunctionInfos()).ToList();
                var xmlInfos = _xmlProvider.GetFunctionInfos();
                return workbookInfos.Concat(xmlInfos).ToList();
            }
        }

        #endregion

        void Excel_WorkbookOpen(Workbook wb)
        {
            var name = wb.Name;
            var regInfo = new WorkbookRegistrationInfo(name);
            lock (_workbookRegistrationInfos)
            {
                _workbookRegistrationInfos[name] = regInfo;
                RegisterWithXmlProvider(wb);
                OnInvalidate();
            }
        }

        void Excel_WorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            // Do we have to worry about renaming / Save As?
            // Do we have to worry about other BeforeClose handlers cancelling the close?
            lock (_workbookRegistrationInfos)
            {
                _workbookRegistrationInfos.Remove(wb.Name);
                UnregisterWithXmlProvider(wb);
            }
        }

        void Excel_WorkbookAddinInstall(Workbook wb)
        {
            var name = wb.Name;
            var regInfo = new WorkbookRegistrationInfo(name);
            lock (_workbookRegistrationInfos)
            {
                _workbookRegistrationInfos[name] = regInfo;
                RegisterWithXmlProvider(wb);
                OnInvalidate();
            }
        }

        void Excel_WorkbookAddinUninstall(Workbook wb)
        {
            lock (_workbookRegistrationInfos)
            {
                _workbookRegistrationInfos.Remove(wb.Name);
                UnregisterWithXmlProvider(wb);
            }
        }

        void RegisterWithXmlProvider(Workbook wb)
        {
            var path = wb.FullName;
            var xmlPath = GetXmlPath(path);
            _xmlProvider.RegisterXmlFunctionInfo(xmlPath);  // Will check if file exists

            var customXmlParts = wb.CustomXMLParts.SelectByNamespace(XmlIntelliSense.Namespace);
            if (customXmlParts.Count > 0)
            {
                // We just take the first one - register against the Bworkbook name
                _xmlProvider.RegisterXmlFunctionInfo(path, customXmlParts[1].XML);
            }
        }

        void UnregisterWithXmlProvider(Workbook wb)
        {
            var path = wb.FullName;
            var xmlPath = GetXmlPath(path);
            _xmlProvider.UnregisterXmlFunctionInfo(path);
            _xmlProvider.UnregisterXmlFunctionInfo(xmlPath);
        }

        void OnInvalidate()
        {
            Invalidate?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            Logger.Provider.Info("WorkbookIntelliSenseProvider.Dispose");
            try
            {
                var xlApp = (Application)ExcelDnaUtil.Application;
                // might fail, e.g. if the process is exiting
                if (xlApp != null)
                {
                    xlApp.WorkbookOpen -= Excel_WorkbookOpen;
                    xlApp.WorkbookBeforeClose -= Excel_WorkbookBeforeClose;
                    xlApp.WorkbookAddinInstall -= Excel_WorkbookAddinInstall;
                    xlApp.WorkbookAddinUninstall -= Excel_WorkbookAddinUninstall;
                }
            }
            catch (Exception ex)
            {
                Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.Dispose Error {ex}");
            }
        }

        string GetXmlPath(string wbPath) => Path.ChangeExtension(wbPath, ".intellisense.xml");
    }
}
