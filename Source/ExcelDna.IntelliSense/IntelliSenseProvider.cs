using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace ExcelDna.IntelliSense
{
    // An IntelliSenseProvider is a part of an IntelliSenseServer that provides IntelliSense info to the server.
    // The providers are built in to the ExcelDna.IntelliSense assembly - there are complications in making this a part that can be extended 
    // by a specific add-in (server activation, cross AppDomain loading etc.).
    // Higher versions of the ExcelDna.IntelliSenseServer are expected to increase the number of providers
    // and/or the scope of some provider (e.g. add support for enums).

    // The server, upon activation and at other times (when? when ExcelDna.IntelliSense.Refresh is called ?) will call the provider to get the IntelliSense info.
    // Maybe the provider can also raise an Invalidate event, to prod the server into reloading the IntelliSense info for that provider.
    // E.g. the XmlProvider might read from a file, and put a FileWatcher on the file so that whenever the file changes, 
    // the server calls back to get the updated info.

    // A major concern is the context in which the provider is called from the server.
    // We separate the Refresh call from the calls to get the info:
    // The Refresh calls are always in a macro context, from the main Excel thread and should be as fast as possible.
    // The GetXXXInfo calls can be made from any thread, should be thread-safe and not call back to Excel.
    // Invalidate can be raised, but the Refresh call might come a lot later...?

    // We expect the server to hook some Excel events to provide the entry points... (not sure what this means anymore...?)

    // Consider interaction with Application.MacroOptions. (or not?)

    interface IIntelliSenseProvider
    {
        void Refresh(); // Executed in a macro context, on the main Excel thread
        IEnumerable<IntelliSenseFunctionInfo> GetFunctionInfos(); // Called from a worker thread - no Excel or COM access (probably an MTA thread involved in the UI Automation)
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

            // Called in a macro context
            public void Refresh()
            {
                if (_regInfoNotAvailable)
                    return;

                object regInfoResponse = ExcelIntegration.GetRegistrationInfo(_xllPath, _version);

                if (regInfoResponse.Equals(ExcelError.ExcelErrorNA))
                {
                    _regInfoNotAvailable = true;
                    return;
                }

                if (regInfoResponse == null)
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

                for (int i = 0; i < regInfo.GetLength(0); i++)
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
                            XllPath = _xllPath
                        };
                    }
                }
            }
        }

        Dictionary<string, XllRegistrationInfo> _xllRegistrationInfos = new Dictionary<string, XllRegistrationInfo>();

        public void Refresh()
        {
            foreach (var xllPath in GetLoadedXllPaths())
            {
                XllRegistrationInfo regInfo;
                if (!_xllRegistrationInfos.TryGetValue(xllPath, out regInfo))
                {
                    regInfo = new XllRegistrationInfo(xllPath);
                    _xllRegistrationInfos[xllPath] = regInfo;
                }
                regInfo.Refresh();
            }
        }

        public IEnumerable<IntelliSenseFunctionInfo> GetFunctionInfos()
        {
            return _xllRegistrationInfos.Values.SelectMany(ri => ri.GetFunctionInfos());
        }

        // Called in macro context
        // Might be implemented by COM AddIns check, or checking loaded Modules with Win32
        // Application.AddIns2 also lists add-ins interactively loaded (Excel 2010+) and has IsOpen property.
        // See: http://blogs.office.com/2010/02/16/migrating-excel-4-macros-to-vba/
        // Alternative on old Excel is DOCUMENTS(2) which lists all loaded .xlm (also .xll?)
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
    }

    // VBA might be easy, since the functions are scoped to the Workbook or add-in, and so might have the right name?
    //         Excel4v(xlcRun, &xlRes, 1, TempStr(" myVBAAddin.xla!myFunc"));

    class VbaIntelliSenseProvider
    {
    }

    // The idea is that other add-in tools like XLW or XLL+ could provide IntelliSense info with an xml file.
    // CONSIDER: How to find these files - can't jsut be relative to the IntelliSense add-in. 
    //           (Maybe next to the foreign .xll file (.xllinfo)?)
    class XmlIntelliSenseProvider
    {
    }
}
