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
    public static class TestFunctions
    {
        [ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
        public static double AddThem(
            [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] double v1,
            [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     double v2)
        {
            return v1 + v2;
        }

        [Description("Test function for the amazing Excel-DNA IntelliSense feature")]
        public static string jDummyFunc()
        {
            return "Howzit !";
        }

    }

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
