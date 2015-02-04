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
        private IntellisenseHelper _intellisense;
        public void AutoOpen()
        {
            _intellisense = new IntellisenseHelper();
        }

        public void AutoClose()
        {
            _intellisense.Dispose();
        }
    }
}
