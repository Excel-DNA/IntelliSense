using System;
using System.Collections.Generic;
using System.Xml.Serialization;

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

        IList<FunctionInfo> GetFunctionInfos(); // Called from a worker thread - no Excel or COM access (probably an MTA thread involved in the UI Automation)
    }
    
    public class FunctionInfo
    {
        public class ArgumentInfo
        {
            public string Name;
            public string Description;
            public string HelpTopic;
        }

        public string Name;
        public string Description;
        public string HelpTopic;
        [XmlElement("Argument", typeof(ArgumentInfo))]
        public List<ArgumentInfo> ArgumentList;
        public string SourcePath; // XllPath for .xll, Workbook Name for Workbook
    }

}
