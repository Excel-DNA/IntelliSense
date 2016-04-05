using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using ExcelDna.IntelliSense;

namespace TestAddIn
{
    public class AddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Register();
            // ExcelDna.Logging.LogDisplay.Show();
        }

        public void AutoClose()
        {
            // CONSIDER: Do we implement an explicit call here, or is the AppDomain Unload event good enough?
        }
    }
}
