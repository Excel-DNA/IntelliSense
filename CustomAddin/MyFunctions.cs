using System;
using ExcelDna.Integration;

namespace ExcelDna.CustomAddin
{
    public class MyFunctions
    {
        [ExcelFunction(Description = "A useful test function that multiply two numbers, and returns the product.")]
        public static double MultiplyThem(
            [ExcelArgument(Name = "Augend", Description = "is the first number, to which will be multiplied")] double v1,
            [ExcelArgument(Name = "Addend", Description = "is the second number that will be multiplied")]     double v2)
        {
            return v1 * v2;
        }

    }
}
