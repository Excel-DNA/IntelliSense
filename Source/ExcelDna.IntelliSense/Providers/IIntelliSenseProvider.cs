using System;
using System.Collections.Generic;
using System.IO;
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

    // Upon activation the server will call the provider to get the IntelliSense info.
    // The provider can also raise an Invalidate event, to prod the server into reloading the IntelliSense info for that provider
    // (a bit like the ribbon Invalidate works).
 
    // A major concern is the context in which the provider is called from the server.
    // We separate the Refresh call from the calls to get the info:
    // The Refresh calls are always in a macro context, from the main Excel thread and should be as fast as possible.
    // The GetXXXInfo calls can be made from any thread, should be thread-safe and not call back to Excel.
    // Invalidate can be raised, but the Refresh call might come a lot later...?

    // We expect the server to hook some Excel events to provide the entry points... (not sure what this means anymore...?)

    // CONSIDER: Is there a way to get the register info from calls to Application.MacroOptions?

    // TODO: We might relax the threading rules, to say that Refresh runs on the same thread as Invalidate 
    // TODO: We might get rid of Refresh (since that runs in the Invalidate context)
    // CONSIDER: The two providers have been refactored to work very similarly - maybe be can extract out a base class...
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
            [XmlAttribute]
            public string Name;
            [XmlAttribute]
            public string Description;
        }

        [XmlAttribute]
        public string Name;
        [XmlAttribute]
        public string Description;
        [XmlAttribute]
        public string HelpTopic;
        [XmlElement("Argument", typeof(ArgumentInfo))]
        public List<ArgumentInfo> ArgumentList;
        [XmlIgnore]
        public string SourcePath; // XllPath for .xll, Workbook Name for Workbook, .xml file path for Xml file

        // Not sure where to put this...
        internal static string ExpandHelpTopic(string path, string helpTopic)
        {
            if (string.IsNullOrEmpty(helpTopic))
                return helpTopic;

            if (helpTopic.StartsWith("http://") || helpTopic.StartsWith("https://") || helpTopic.StartsWith("file://"))
            {
                if (helpTopic.EndsWith("!0"))
                {
                    helpTopic = helpTopic.Substring(0, helpTopic.Length - 2);
                }
            }
            else if (!Path.IsPathRooted(helpTopic))
            {
                helpTopic = Path.Combine(path, helpTopic);
            }
            return helpTopic;
        }


    }

}
