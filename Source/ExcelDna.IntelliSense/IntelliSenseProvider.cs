using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;

namespace ExcelDna.IntelliSense
{
    // An IntelliSenseProvider is the part of an IntelliSenseServer that provides the IntelliSense info gathered from add-ins to the server.
    // These providers are built in to the ExcelDna.IntelliSense assembly - there are complications in making this a part that can be extended 
    // by a specific add-in (server activation, cross AppDomain loading etc.).
    //
    // Higher versions of the ExcelDna.IntelliSenseServer are expected to increase the number of providers
    // and/or the scope of some provider (e.g. add support for enums).
    // No provision is made at the moment for user-created providers or an external provider API.

    // The server, upon activation and at other times (when? when ExcelDna.IntelliSense.Refresh is called ?) will call the provider to get the IntelliSense info.
    // Maybe the provider can also raise an Invalidate event, to prod the server into reloading the IntelliSense info for that provider
    // (a bit like the ribbon Invalidate works).
    // E.g. the XmlProvider might read from a file, and put a FileWatcher on the file so that whenever the file changes, 
    // the server calls back to get the updated info.

    // A major concern is the context in which the provider is called from the server.
    // We separate the Refresh call from the calls to get the info:
    // The Refresh calls are always in a macro context, from the main Excel thread and should be as fast as possible.
    // The GetXXXInfo calls can be made from any thread, should be thread-safe and not call back to Excel.
    // Invalidate can be raised, but the Refresh call might come a lot later...?

    // We expect the server to hook some Excel events to provide the entry points... (not sure what this means anymore...?)

    // TODO: Consider interaction with Application.MacroOptions. (or not?)

    // TODO: We might relax the threading rules, to say that Refresh runs on the same thread as Invalidate 
    // TODO: We might get rid of Refresh (since that runs in the Invalidate context)
    // TODO: The two providers have been refactored to work very similarly - maybe be can extract out a base class...
    interface IIntelliSenseProvider : IDisposable
    {
        void Initialize();  // Executed in a macro context, on the main Excel thread
        void Refresh();     // Executed in a macro context, on the main Excel thread
        event EventHandler Invalidate;  // Must be raised in a macro context, on the main thread

        IList<IntelliSenseFunctionInfo> GetFunctionInfos(); // Called from a worker thread - no Excel or COM access (probably an MTA thread involved in the UI Automation)
    }

    // Provides IntelliSense info for all Excel-DNA based .xll add-ins, using the built-in RegistrationInfo helper function.
    class ExcelDnaIntelliSenseProvider : IIntelliSenseProvider
    {
        class XllRegistrationInfo
        {
            readonly string _xllPath;
            bool _regInfoNotAvailable = false;  // Set to true if we know for sure that reginfo is #N/A
            double _version = -1;               // Version indicator to enumerate from scratch
            object[,] _regInfo = null;          // Default value

            public XllRegistrationInfo(string xllPath)
            {
                _xllPath = xllPath;
            }

            // Called on the main thread in a macro context
            public void Refresh()
            {
                if (_regInfoNotAvailable)
                    return;

                object regInfoResponse = ExcelIntegration.GetRegistrationInfo(_xllPath, _version);

                if (regInfoResponse.Equals(ExcelError.ExcelErrorNA))
                {
                    _regInfoNotAvailable = true;
                    Logger.Provider.Verbose($"XllRegistrationInfo not available for {_xllPath}");
                    return;
                }

                if (regInfoResponse == null || regInfoResponse.Equals(ExcelError.ExcelErrorNum))
                {
                    // no update - versions match
                    return;
                }

                _regInfo = regInfoResponse as object[,];
                if (_regInfo != null)
                {
                    Debug.Assert((string)_regInfo[0, 0] == _xllPath);
                    _version = (double)_regInfo[0, 1];
                }
            }

            // Not in macro context - don't call Excel, could be any thread.
            public IEnumerable<IntelliSenseFunctionInfo> GetFunctionInfos()
            {
                // to avoid worries about locking and this being updated from another thread, we take a copy of _regInfo
                var regInfo = _regInfo;
                /*
                    result[0, 0] = XlAddIn.PathXll;
                    result[0, 1] = registrationInfoVersion;
                 */

                if (regInfo == null)
                    yield break;

                int regInfoCount = regInfo.GetLength(0);
                Logger.Provider.Verbose($"XllRegistrationInfo for {_xllPath}: {regInfoCount} registrations");
                for (int i = 0; i < regInfoCount; i++)
                {
                    if (regInfo[i, 0] is ExcelEmpty)
                    {
                        string functionName = regInfo[i, 3] as string;
                        string description = regInfo[i, 9] as string;

                        string argumentStr = regInfo[i, 4] as string;
                        string[] argumentNames = string.IsNullOrEmpty(argumentStr) ? new string[0] : argumentStr.Split(',');

                        List<IntelliSenseFunctionInfo.ArgumentInfo> argumentInfos = new List<IntelliSenseFunctionInfo.ArgumentInfo>();
                        for (int j = 0; j < argumentNames.Length; j++)
                        {
                            argumentInfos.Add(new IntelliSenseFunctionInfo.ArgumentInfo 
                            { 
                                ArgumentName = argumentNames[j], 
                                Description = regInfo[i, j + 10] as string 
                            });
                        }

                        yield return new IntelliSenseFunctionInfo
                        {
                            FunctionName = functionName,
                            Description = description,
                            ArgumentList = argumentInfos,
                            SourcePath = _xllPath
                        };
                    }
                }
            }
        }

        ExcelSynchronizationContext _syncContextExcel;
        Dictionary<string, XllRegistrationInfo> _xllRegistrationInfos = new Dictionary<string, XllRegistrationInfo>();
        LoaderNotification _loaderNotification;
        bool _isDirty;
        public event EventHandler Invalidate;

        public ExcelDnaIntelliSenseProvider()
        {
            _loaderNotification = new LoaderNotification();
            _loaderNotification.LoadNotification += loaderNotification_LoadNotification;
            _syncContextExcel = new ExcelSynchronizationContext();
        }

        #region IIntelliSenseProvider implementation

        // Must be called on the main Excel thread
        public void Initialize()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Initialize");
            lock (_xllRegistrationInfos)
            {
                foreach (var xllPath in GetLoadedXllPaths())
                {
                    if (!_xllRegistrationInfos.ContainsKey(xllPath))
                    {
                        Logger.Provider.Verbose($"ExcelDnaIntelliSenseProvider.Initialize: Adding XllRegistrationInfo for {xllPath}");
                        XllRegistrationInfo regInfo = new XllRegistrationInfo(xllPath);
                        _xllRegistrationInfos[xllPath] = regInfo;
                        regInfo.Refresh();
                    }
                }
            }
        }

        // Must be called on the main Excel thread
        public void Refresh()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Refresh");
            lock (_xllRegistrationInfos)
            {
                foreach (var regInfo in _xllRegistrationInfos.Values)
                {
                    regInfo.Refresh();
                }
                _isDirty = false;
            }
        }

        // May be called from any thread
        public IList<IntelliSenseFunctionInfo> GetFunctionInfos()
        {
            IList<IntelliSenseFunctionInfo> functionInfos;
            lock (_xllRegistrationInfos)
            {
                functionInfos = _xllRegistrationInfos.Values.SelectMany(ri => ri.GetFunctionInfos()).ToList();
            }
            Logger.Provider.Verbose("ExcelDnaIntelliSenseProvider.GetFunctionInfos Begin");
            foreach (var info in functionInfos)
            {
                Logger.Provider.Verbose($"\t{info.FunctionName}({info.ArgumentList.Count}) - {info.Description} ");
            }
            Logger.Provider.Verbose("ExcelDnaIntelliSenseProvider.GetFunctionInfos End");
            return functionInfos;
        }

        #endregion

        // DANGER: Still subject to LoaderLock problem...
        // TODO: Consider Load/Unload done when AddIns is enumerated...
        void loaderNotification_LoadNotification(object sender, LoaderNotification.NotificationEventArgs e)
        {
            // Debug.Print($"@>@>@>@> LoadNotification: {e.Reason} - {e.FullDllName}");
            if (e.FullDllName.EndsWith(".xll", StringComparison.OrdinalIgnoreCase))
                _syncContextExcel.Post(ProcessLoadNotification, e);
        }

        // Runs on the main thread, in a macro context 
        void ProcessLoadNotification(object state)
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            // we might want to introduce a delay here, so that the .xll can complete loading...
            var notification = (LoaderNotification.NotificationEventArgs)state;;
            var xllPath = notification.FullDllName;

            Logger.Provider.Verbose($"ExcelDnaIntelliSenseProvider.ProcessLoadNotification {notification}, {xllPath}");

            lock (_xllRegistrationInfos)
            {
                XllRegistrationInfo regInfo;
                if (!_xllRegistrationInfos.TryGetValue(xllPath, out regInfo))
                {
                    if (notification.Reason == LoaderNotification.Reason.Loaded)
                    {
                        regInfo = new XllRegistrationInfo(xllPath);
                        _xllRegistrationInfos[xllPath] = regInfo;
                        //regInfo.Refresh();    // Rather not.... so that we don't even try during the AddIns enumeration... OnInvalidate will lead to Refresh()

                        if (!_isDirty)
                        {
                            _isDirty = true;
                            _syncContextExcel.Post(OnInvalidate, null);
                        }
                    }
                }
                else if (notification.Reason == LoaderNotification.Reason.Unloaded)
                {
                    _xllRegistrationInfos.Remove(xllPath);
                    // Not too worried about cleaning up
                    // OnInvalidate();
                }
            }
        }

        // Called in macro context
        // Might be implemented by COM AddIns check, or checking loaded Modules with Win32
        // Application.AddIns2 also lists add-ins interactively loaded (Excel 2010+) and has IsOpen property.
        // See: http://blogs.office.com/2010/02/16/migrating-excel-4-macros-to-vba/
        // Alternative on old Excel is DOCUMENTS(2) which lists all loaded .xlm (also .xll?)
        // Alternative, more in line with our update watch, is to enumerate all loaded modules...
        IEnumerable<string> GetLoadedXllPaths()
        {
            // TODO: Implement properly...
            dynamic app = ExcelDnaUtil.Application;
            foreach (var addin in app.AddIns2)
            {
                if (addin.IsOpen && Path.GetExtension(addin.FullName) == ".xll")
                {
                    yield return addin.FullName;
                }
            }
        }

        void OnInvalidate(object _unused_)
        {
            Invalidate?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
            _loaderNotification.Dispose();
        }
    }

    // For VBA code, (either in a regular workbook that is open, or in an add-in)
    // we allow the IntelliSense info to be put into a worksheet (possibly hidden or very hidden)
    // In the Workbook that contains the VBA.
    // Initially we won't scope the IntelliSense to the Workbook where the UDFs are defined, 
    // but we should consider that.

    // TODO: Can't we read the Application.MacroOptions...?

    class WorkbookIntelliSenseProvider : IIntelliSenseProvider
    {
        const string intelliSenseWorksheetName = "_IntelliSense_FunctionInfo_";
        class WorkbookRegistrationInfo
        {
            readonly string _name;
            DateTime _lastUpdate;               // Version indicator to enumerate from scratch
            object[,] _regInfo = null;          // Default value

            public WorkbookRegistrationInfo(string name)
            {
                _name = name;
            }
 
            // Called in a macro context
            public void Refresh()
            {
                dynamic app = ExcelDnaUtil.Application;
                var wb = app.Workbooks[_name];

                try
                {
                    var ws = wb.Sheets[intelliSenseWorksheetName];
                    var info = ws.UsedRange.Value;
                    _regInfo = info as object[,];
                }
                catch (Exception ex)
                {
                    // We expect this if there is no sheet.
                    // Another approach would be xlSheetNm
                    Debug.Print("WorkbookIntelliSenseProvider.Refresh Error : " + ex.Message);
                    _regInfo = null;
                }
                _lastUpdate = DateTime.Now;
            }

            // Not in macro context - don't call Excel, could be any thread.
            public IEnumerable<IntelliSenseFunctionInfo> GetFunctionInfos()
            {
                // to avoid worries about locking and this being updated from another thread, we take a copy of _regInfo
                var regInfo = _regInfo;
                /*
                    result[0, 0] = XlAddIn.PathXll;
                    result[0, 1] = registrationInfoVersion;
                 */

                if (regInfo == null)
                    yield break;

                // regInfo is 1-based: object[1..x, 1..y].
                for (int i = 1; i <= regInfo.GetLength(0); i++)
                {
                    string functionName = regInfo[i, 1] as string;
                    string description = regInfo[i, 2] as string;

                    List<IntelliSenseFunctionInfo.ArgumentInfo> argumentInfos = new List<IntelliSenseFunctionInfo.ArgumentInfo>();
                    for (int j = 3; j <= regInfo.GetLength(1) - 1; j += 2)
                    {
                        var arg = regInfo[i, j] as string;
                        var argDesc = regInfo[i, j + 1] as string;
                        if (!string.IsNullOrEmpty(arg))
                        {
                            argumentInfos.Add(new IntelliSenseFunctionInfo.ArgumentInfo
                            {
                                ArgumentName = arg,
                                Description = argDesc
                            });
                        }
                    }

                    yield return new IntelliSenseFunctionInfo
                    {
                        FunctionName = functionName,
                        Description = description,
                        ArgumentList = argumentInfos,
                        SourcePath = _name
                    };
                }
            }
        }

        Dictionary<string, WorkbookRegistrationInfo> _workbookRegistrationInfos = new Dictionary<string, WorkbookRegistrationInfo>();

        public event EventHandler Invalidate;

        #region IIntelliSenseProvider implementation

        public WorkbookIntelliSenseProvider()
        {
        }

        public void Initialize()
        {
            Logger.Provider.Info("WorkbookIntelliSenseProvider.Initialize");

            // The events are just to keep track of the set of open workbooks, 
            var xlApp = (Application)ExcelDnaUtil.Application;
            xlApp.WorkbookOpen += Excel_WorkbookOpen;
            xlApp.WorkbookBeforeClose += Excel_WorkbookBeforeClose;
            //xlApp.WorkbookAddinInstall += Excel_WorkbookAddinInstall;
            //xlApp.WorkbookAddinUninstall += Excel_WorkbookAddinUninstall;

            //var app = ExcelDnaUtil.Application;
            // app.WorkbookLoaded...
            lock (_workbookRegistrationInfos)
            {
                foreach (var name in GetLoadedWorkbookNames())
                {
                    if (!_workbookRegistrationInfos.ContainsKey(name))
                    {
                        WorkbookRegistrationInfo regInfo = new WorkbookRegistrationInfo(name);
                        _workbookRegistrationInfos[name] = regInfo;
                        regInfo.Refresh();
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
            }
        }

        // May be called from any thread
        public IList<IntelliSenseFunctionInfo> GetFunctionInfos()
        {
            lock (_workbookRegistrationInfos)
            {
                return _workbookRegistrationInfos.Values.SelectMany(ri => ri.GetFunctionInfos()).ToList();
            }
        }

        #endregion

        void Excel_WorkbookOpen(Workbook Wb)
        {
            var regInfo = new WorkbookRegistrationInfo(Wb.Name);
            lock (_workbookRegistrationInfos)
            {
                _workbookRegistrationInfos[Wb.Name] = regInfo;
                OnInvalidate();
            }
        }

        void Excel_WorkbookBeforeClose(Workbook Wb, ref bool Cancel)
        {
            // Do we have to worry about renaming / Save As?
            lock (_workbookRegistrationInfos)
            {
                _workbookRegistrationInfos.Remove(Wb.Name);
            }
        }

        //private void Excel_WorkbookAddinInstall(Workbook Wb)
        //{
        //    throw new NotImplementedException();
        //}

        //private void Excel_WorkbookAddinUninstall(Workbook Wb)
        //{
        //    throw new NotImplementedException();
        //}

        // Called in macro context
        // Might be implemented by tracking Application events
        // Remember this changes when a workbook is saved, and can refer to the wrong workbook as they are closed / opened
        // CONSIDER: Check AddIns2 ?
        IEnumerable<string> GetLoadedWorkbookNames()
        {
            // TODO: Implement properly...
            dynamic app = ExcelDnaUtil.Application;
            foreach (var wb in app.Workbooks)
            {
                yield return wb.Name; 
            }
        }

        void OnInvalidate()
        {
            Invalidate?.Invoke(this, EventArgs.Empty);
        }

        public void Dispose()
        {
        }
    }

    // The idea is that other add-in tools like XLW or XLL+ could provide IntelliSense info with an xml file.
    // CONSIDER: How to find these files - can't just be relative to the IntelliSense add-in. 
    //           (Maybe next to the foreign .xll file (.xllinfo)?)
    class XmlIntelliSenseProvider
    {
    }
}
