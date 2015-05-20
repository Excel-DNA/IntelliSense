using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelDna.IntelliSense
{
    // An IntelliSenseProvider provides IntelliSense info to the server.
    // The providers are built in to the ExcelDna.IntelliSense assembly - there are complications in making this a part that can be extended 
    // by a specific add-in (server activation, cross AppDomain loading etc.). Higher versions of the ExcelDna.IntelliSenseServer are expected to increase the number of providers
    // and/or the scope of some provider (e.g. add support for enums).

    // The server, upon activation and at other times (when?) will call the provider to get the IntelliSense info.
    // The provider can also raise an Invalidate event, to prod the server into reloading the IntelliSense info for that provider.
    // E.g. the XmlProvider might read from a file, and put a FileWatcher on the file so that whenever the file changes, 
    // the server calls back to get the updated info.

    // A major concern is the context in which the provider is called from the server.
    // We separate the Update call from the calls to get the info:
    // The Update calls are always in a macro context, from the main Excel thread and should be as fast as possible.
    // Maybe Update returns a Task<Info> but must do its work on the main thread before returning?
    // The GetXXXInfo calls can be made from any thread, should be thread-safe and not call back to Excel.
    // Invalidate can be raised, but the update call might come a lot later.
    // We expect the server to hhok some Excel events to provide the entry points...
    interface IIntelliSenseProvider
    {

    }

    class XllIntelliSenseProvider
    {
    }

    // VBA might be easy, since the functions are scoped to the Workbook or add-in, and so might have the right name?
    class VbaIntelliSenseProvider
    {
    }

    class XmlIntelliSenseProvider
    {
    }
}
