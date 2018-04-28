#if DEBUG
using System;
using System.ComponentModel;
using ExcelDna.Integration;
using ExcelDna.Logging;

namespace ExcelDna.CustomAddin
    {
        // These functions - are just here for testing...
        public class MyFunctions
        {
            [ExcelFunction(Description = "Returns the sum of two particular numbers that are given\r\n(As a test, of course)",
                           HelpTopic = "http://www.google.com")]
            public static double AddThem(
                [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] double v1,
                [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     double v2)
            {
                return v1 + v2;
            }

            [ExcelFunction(Description = "--------------------",
                           HelpTopic = "MyFile.chm!100")]
            public static double AdxThem(
                [ExcelArgument(Name = "[tag]", Description = "is the first number, to which will be added")] double v1,
                [ExcelArgument(Name = "[Addend]", Description = "is the second number that will be added")]     double v2)
            {
                return v1 + v2;
            }

            [Description("Test function for the amazing Excel-DNA IntelliSense feature")]
            public static string jDummyFunc()
            {
                return "Howzit !";
            }

            [ExcelFunction(Name ="a.test?d_.3")]
            public static object AnotherFunction( [Description("In and out")] object inout)
            {
                return inout;
            }

            [ExcelFunction(Name ="A.Non.Descript.Function")]
            public static object ANonDescriptFunction(object inout)
            {
                return inout;
            }

            [ExcelFunction(Name ="A.Descript.Function", Description = "Has a description")]
            public static object ADescriptFunction(object inout)
            {
                return inout;
            }

            [ExcelFunction(Description = "Has many arguments")]
            public static object AManyArgFunction(
                [ExcelArgument(Name = "Argument1", Description = "is the first argument")] double arg1,
                [ExcelArgument(Name = "Argument2", Description = "is another argument")] double arg2,
                [ExcelArgument(Name = "Argument3", Description = "is another argument")] double arg3,
                [ExcelArgument(Name = "Argument4", Description = "is another argument")] double arg4,
                [ExcelArgument(Name = "Argument5", Description = "is another argument")] double arg5,
                [ExcelArgument(Name = "Argument6", Description = "is another argument")] double arg6,
                [ExcelArgument(Name = "Argument7", Description = "is another argument")] double arg7,
                [ExcelArgument(Name = "Argument8", Description = "is another argument")] double arg8,
                [ExcelArgument(Name = "Argument9", Description = "is another argument")] double arg9,
                [ExcelArgument(Name = "Argument10", Description = "is another argument")] double arg10,
                [ExcelArgument(Name = "Argument11", Description = "is another argument")] double arg11,
                [ExcelArgument(Name = "Argument12", Description = "is another argument")] double arg12,
                [ExcelArgument(Name = "Argument13", Description = "is another argument")] double arg13,
                [ExcelArgument(Name = "Argument14", Description = "is another argument")] double arg14,
                [ExcelArgument(Name = "Argument15", Description = "is another argument")] double arg15,
                [ExcelArgument(Name = "Argument16", Description = "is another argument")] double arg16,
                [ExcelArgument(Name = "Argument18", Description = "is another argument")] double arg18,
                [ExcelArgument(Name = "Argument19", Description = "is another argument")] double arg19,
                [ExcelArgument(Name = "Argument20", Description = "is another argument")] double arg20
                )
            {
                return arg1;
            }

        [ExcelCommand]
            public static void dnaLogDisplayShow()
            {
                LogDisplay.Show();
            }
        }
    }
#endif
