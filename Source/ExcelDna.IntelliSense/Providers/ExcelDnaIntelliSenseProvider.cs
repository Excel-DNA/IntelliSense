using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelDna.Integration;

namespace ExcelDna.IntelliSense
{
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
                    Logger.Provider.Info($"XllRegistrationInfo not available for {_xllPath}");
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
            public IEnumerable<FunctionInfo> GetFunctionInfos()
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
                        string helpTopic = regInfo[i, 8] as string;
                        string description = regInfo[i, 9] as string;

                        string argumentStr = regInfo[i, 4] as string;
                        string[] argumentNames = string.IsNullOrEmpty(argumentStr) ? new string[0] : argumentStr.Split(',');

                        List<FunctionInfo.ArgumentInfo> argumentInfos = new List<FunctionInfo.ArgumentInfo>();
                        for (int j = 0; j < argumentNames.Length; j++)
                        {
                            argumentInfos.Add(new FunctionInfo.ArgumentInfo 
                            { 
                                Name = argumentNames[j], 
                                Description = regInfo[i, j + 10] as string 
                            });
                        }

                        yield return new FunctionInfo
                        {
                            Name = functionName,
                            Description = description,
                            HelpTopic = helpTopic,
                            ArgumentList = argumentInfos,
                            SourcePath = _xllPath
                        };
                    }
                }
            }
        }

        ExcelSynchronizationContext _syncContextExcel;
        XmlIntelliSenseProvider _xmlProvider;
        Dictionary<string, XllRegistrationInfo> _xllRegistrationInfos = new Dictionary<string, XllRegistrationInfo>();
        LoaderNotification _loaderNotification;
        bool _isDirty;
        public event EventHandler Invalidate;

        public ExcelDnaIntelliSenseProvider()
        {
            _loaderNotification = new LoaderNotification();
            _loaderNotification.LoadNotification += loaderNotification_LoadNotification;
            _syncContextExcel = new ExcelSynchronizationContext();
            _xmlProvider = new XmlIntelliSenseProvider();
            _xmlProvider.Invalidate += (sender, e) => OnInvalidate(null);
        }

        #region IIntelliSenseProvider implementation

        // Must be called on the main Excel thread
        public void Initialize()
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Initialize");

            _xmlProvider.Initialize();

            lock (_xllRegistrationInfos)
            {
                foreach (var xllPath in GetLoadedXllPaths())
                {
                    if (!_xllRegistrationInfos.ContainsKey(xllPath))
                    {
                        Logger.Provider.Verbose($"ExcelDnaIntelliSenseProvider.Initialize: Adding XllRegistrationInfo for {xllPath}");
                        XllRegistrationInfo regInfo = new XllRegistrationInfo(xllPath);
                        _xllRegistrationInfos[xllPath] = regInfo;

                        _xmlProvider.RegisterXmlFunctionInfo(GetXmlPath(xllPath));
                        
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
                _xmlProvider.Refresh();
                _isDirty = false;
            }
        }

        // May be called from any thread
        public IList<FunctionInfo> GetFunctionInfos()
        {
            IList<FunctionInfo> excelDnaInfos;
            lock (_xllRegistrationInfos)
            {
                excelDnaInfos = _xllRegistrationInfos.Values.SelectMany(ri => ri.GetFunctionInfos()).ToList();
            }
            Logger.Provider.Verbose("ExcelDnaIntelliSenseProvider.GetFunctionInfos Begin");
            foreach (var info in excelDnaInfos)
            {
                Logger.Provider.Verbose($"\t{info.Name}({info.ArgumentList.Count}) - {info.Description} ");
            }

            var xmlInfos = _xmlProvider.GetFunctionInfos();
            var allInfos = excelDnaInfos.Concat(xmlInfos).ToList();

            Logger.Provider.Verbose("ExcelDnaIntelliSenseProvider.GetFunctionInfos End");
            return allInfos;
        }

        #endregion

        // DANGER: Still subject to LoaderLock warning ...
        // TODO: Consider Load/Unload done when AddIns is enumerated...
        void loaderNotification_LoadNotification(object sender, LoaderNotification.NotificationEventArgs e)
        {
            // Debug.Print($"@>@>@>@> LoadNotification: {e.Reason} - {e.FullDllName}");
            if (e.FullDllName.EndsWith(".xll", StringComparison.OrdinalIgnoreCase))
                _syncContextExcel.Post(ProcessLoadNotification, e);
        }

        // Runs on the main thread, in a macro context 
        // WARNING: The sequence of calls here, between queued 
        //          instances of ProcessLoadNotification, Refresh and OnInvalidate
        //          is quite fragile.
        void ProcessLoadNotification(object state)
        {
            Debug.Assert(Thread.CurrentThread.ManagedThreadId == 1);
            // we might want to introduce a delay here, so that the .xll can complete loading...
            var notification = (LoaderNotification.NotificationEventArgs)state;
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

                        _xmlProvider.RegisterXmlFunctionInfo(GetXmlPath(xllPath));

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
                    _xmlProvider.UnregisterXmlFunctionInfo(GetXmlPath(xllPath));

                    // Not too eager when cleaning up
                    // OnInvalidate();
                }
            }
        }

        string GetXmlPath(string xllPath) => Path.ChangeExtension(xllPath, ".intellisense.xml");

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
            Logger.Provider.Info("ExcelDnaIntelliSenseProvider.Dispose");
            _loaderNotification.Dispose();
        }
    }
}
