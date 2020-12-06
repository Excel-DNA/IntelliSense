﻿using System;
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

                Logger.Provider.Verbose($"WorkbookRegistrationInfo.Refresh - Workbook {_name} at {_path}");
                try
                {
                    var ws = wb.Sheets[intelliSenseWorksheetName];
                    Logger.Provider.Verbose($"WorkbookRegistrationInfo.Refresh - IntelliSense sheet found");
                    var info = ws.UsedRange.Value;
                    _regInfo = info as object[,];
                    Logger.Provider.Verbose($"WorkbookRegistrationInfo.Refresh - Read {_regInfo.GetLength(0) - 1} registrations");
                }
                catch (Exception ex)
                {
                    // We expect this if there is no sheet.
                    Debug.Print("WorkbookIntelliSenseProvider.Refresh Error : " + ex.Message);
                    Logger.Provider.Verbose($"WorkbookRegistrationInfo.Refresh - No IntelliSense sheet found");
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
                if (numRows < 2 || numCols < 2) // Either no registrations or no descriptions - nothing to register
                    yield break;

                var idCheck = regInfo[1, 1] as string;
                if (!functionInfoId.Equals(idCheck, StringComparison.OrdinalIgnoreCase))
                {
                    Logger.Provider.Warn($"WorkbookIntelliSenseProvider - Invalid FunctionInfo Identifier: ({idCheck})");
                    yield break;
                }

                // regInfo is 1-based: object[1..x, 1..y].
                for (int i = 2; i <= numRows; i++)
                {
                    string functionName = regInfo[i, 1] as string;
                    string description = regInfo[i, 2] as string;
                    string helpTopic = (numCols >= 3) ? (regInfo[i, 3] as string) : "";

                    if (string.IsNullOrEmpty(functionName))
                        continue;

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

                    // Some cleanup and normalization
                    functionName = functionName.Trim();
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
        ExcelSynchronizationContext _syncContextExcel;
        XmlIntelliSenseProvider _xmlProvider;
        public event EventHandler Invalidate;

        #region IIntelliSenseProvider implementation

        public WorkbookIntelliSenseProvider()
        {
            _syncContextExcel = new ExcelSynchronizationContext();
            _xmlProvider = new XmlIntelliSenseProvider();
            _xmlProvider.Invalidate += ( sender, e) => OnInvalidate();
        }

        public void Initialize()
        {
            Logger.Provider.Info("WorkbookIntelliSenseProvider.Initialize");
            _xmlProvider.Initialize();
            _syncContextExcel.Post(OnInitialize, null);
        }

        // Must be called on the main thread, in a macro context
        void OnInitialize(object _unused_)
        {
            Logger.Provider.Info("WorkbookIntelliSenseProvider.OnInitialize");

            // The events are just to keep track of the set of open workbooks, 
            var xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.WorkbookOpen += Excel_WorkbookOpen;
            xlApp.WorkbookBeforeClose += Excel_WorkbookBeforeClose;
            xlApp.WorkbookAddinInstall += Excel_WorkbookAddinInstall;
            xlApp.WorkbookAddinUninstall += Excel_WorkbookAddinUninstall;
            Logger.Provider.Verbose("WorkbookIntelliSenseProvider.OnInitialize - Installed event listeners");

            lock (_workbookRegistrationInfos)
            {
                Logger.Provider.Verbose("WorkbookIntelliSenseProvider.OnInitialize - Starting Workbooks loop");
                foreach (Workbook wb in xlApp.Workbooks)
                {
                    var name = wb.Name;
                    Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.OnInitialize - Adding registration for {name}");
                    if (!_workbookRegistrationInfos.ContainsKey(name))
                    {
                        WorkbookRegistrationInfo regInfo = new WorkbookRegistrationInfo(name);
                        _workbookRegistrationInfos[name] = regInfo;

                        regInfo.Refresh();

                        RegisterWithXmlProvider(wb);
                    }
                }

                // NOTE: This access to AddIns2 might have caused long load delays
                //if (ExcelDnaUtil.ExcelVersion >= 14.0)
                //{
                //    foreach (AddIn addIn in xlApp.AddIns2)
                //    {
                //        if (addIn.IsOpen && Path.GetExtension(addIn.FullName) != ".xll")
                //        {
                //            // Can it be "Open" and not be loaded?
                //            var name = addIn.Name;
                //            Workbook wbAddIn;
                //            try
                //            {
                //                // TODO: Log
                //                wbAddIn = xlApp.Workbooks[name];
                //            }
                //            catch
                //            {
                //                // TODO: Log
                //                continue;
                //            }

                //            WorkbookRegistrationInfo regInfo = new WorkbookRegistrationInfo(name);
                //            _workbookRegistrationInfos[name] = regInfo;

                //            regInfo.Refresh();

                //            RegisterWithXmlProvider(wbAddIn);
                //        }
                //    }
                //}

                Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.OnInitialize - Checking Add-Ins");

                var loadedAddIns = Integration.XlCall.Excel(Integration.XlCall.xlfDocuments, 2) as object[,];
                if (loadedAddIns == null)
                {
                    // This is normal if there are none
                    Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.Initialize - DOCUMENTS(2) returned null");
                    return;
                }
                for (int i = 0; i < loadedAddIns.GetLength(1); i++)
                {
                    var addInName = loadedAddIns[0, i] as string;
                    Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.Initialize - Checking Add-In {addInName}");
                    if (addInName != null && Path.GetExtension(addInName) != ".xll")    // We don't actually expect the .xll add-ins here - and they're taken care of elsewhere
                    {
                        // Can it be "Open" and not be loaded?
                        var name = addInName;
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

                        Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.Initialize - Adding registration for add-in {name}");

                        WorkbookRegistrationInfo regInfo = new WorkbookRegistrationInfo(name);
                        _workbookRegistrationInfos[name] = regInfo;

                        regInfo.Refresh();

                        RegisterWithXmlProvider(wbAddIn);
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
            try
            {
                var regInfo = new WorkbookRegistrationInfo(name);
                lock (_workbookRegistrationInfos)
                {
                    _workbookRegistrationInfos[name] = regInfo;
                    RegisterWithXmlProvider(wb);
                    OnInvalidate();
                }
            }
            catch (Exception ex)
            {
                Logger.Provider.Error(ex, $"Unhandled exception in {nameof(Excel_WorkbookOpen)}, Workbook: {name}");
            }
        }

        void Excel_WorkbookBeforeClose(Workbook wb, ref bool cancel)
        {
            var name = wb.Name;
            try
            {
                // Do we have to worry about renaming / Save As?
                // Do we have to worry about other BeforeClose handlers cancelling the close?
                lock (_workbookRegistrationInfos)
                {
                    _workbookRegistrationInfos.Remove(name);
                    UnregisterWithXmlProvider(wb);
                }
            }
            catch (Exception ex)
            {
                Logger.Provider.Error(ex, $"Unhandled exception in {nameof(Excel_WorkbookBeforeClose)}, Workbook: {name}");
            }

        }

        void Excel_WorkbookAddinInstall(Workbook wb)
        {
            var name = wb.Name;
            try
            {
                var regInfo = new WorkbookRegistrationInfo(name);
                lock (_workbookRegistrationInfos)
                {
                    _workbookRegistrationInfos[name] = regInfo;
                    RegisterWithXmlProvider(wb);
                    OnInvalidate();
                }
            }
            catch (Exception ex)
            {
                Logger.Provider.Error(ex, $"Unhandled exception in {nameof(Excel_WorkbookAddinInstall)}, Workbook: {name}");
            }
        }

        void Excel_WorkbookAddinUninstall(Workbook wb)
        {
            var name = wb.Name;
            try
            {
                lock (_workbookRegistrationInfos)
                {
                    _workbookRegistrationInfos.Remove(wb.Name);
                    UnregisterWithXmlProvider(wb);
                }
            }
            catch (Exception ex)
            {
                Logger.Provider.Error(ex, $"Unhandled exception in {nameof(Excel_WorkbookAddinUninstall)}, Workbook: {name}");
            }
        }

        void RegisterWithXmlProvider(Workbook wb)
        {
            Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.RegisterWithXmlProvider");

            var path = wb.FullName;
            var xmlPath = GetXmlPath(path);
            _xmlProvider.RegisterXmlFunctionInfo(xmlPath);  // Will check if file exists

            Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.RegisterWithXmlProvider - Checking CustomXMLParts");

            var customXmlParts = wb.CustomXMLParts.SelectByNamespace(XmlIntelliSense.Namespace);
            if (customXmlParts.Count > 0)
            {
                Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.RegisterWithXmlProvider - CustomXMLPart found");
                // We just take the first one - register against the Bworkbook name
                _xmlProvider.RegisterXmlFunctionInfo(path, customXmlParts[1].XML);
            }
            else
            {
                Logger.Provider.Verbose($"WorkbookIntelliSenseProvider.RegisterWithXmlProvider - No CustomXMLPart found");
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
