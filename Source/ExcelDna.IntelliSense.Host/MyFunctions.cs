#if DEBUG
    using System;
    using System.ComponentModel;
    using ExcelDna.Integration;

    namespace ExcelDna.CustomAddin
    {
        // These functions - are just here for testing...
        public class MyFunctions
        {
            [ExcelFunction(Description = "A useful test function that adds two numbers, and returns the product.")]
            public static double AddThem(
                [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be multiplied")] double v1,
                [ExcelArgument(Name = "Addend", Description = "is the second number that will be multiplied")]     double v2)
            {
                return v1 + v2;
            }

            [Description("Test function for the amazing Excel-DNA IntelliSense feature")]
            public static string jDummyFunc()
            {
                return "Howzit !";
            }

            [ExcelFunction]
            public static object AnotherFunction( [Description("In and out")] object inout)
            {
                return inout;
            }

            [ExcelFunction(Name ="A.Non.Descript.Function")]
            public static object ANonDescriptFunction(object inout)
            {
                return inout;
            }

        }
    }
#endif
