#if DEBUG
    using System;
    using System.ComponentModel;
    using ExcelDna.Integration;

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
                [ExcelArgument(Name = "tag", Description = "is the first number, to which will be added")] double v1,
                [ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     double v2)
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

        }
    }
#endif
