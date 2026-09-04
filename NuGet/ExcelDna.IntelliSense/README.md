# ExcelDna.IntelliSense

ExcelDna.IntelliSense adds in-sheet function help for Excel user-defined functions. It displays function descriptions and argument help while users type formulas, using the metadata supplied by `ExcelFunction` and `ExcelArgument` attributes.

```csharp
[ExcelFunction(Description = "Adds two numbers.")]
public static double AddThem(
    [ExcelArgument(Description = "The first number.")] double x,
    [ExcelArgument(Description = "The second number.")] double y)
{
    return x + y;
}
```

Install this package when you want IntelliSense support as part of a managed Excel-DNA add-in. For VBA workbooks or for add-ins where you prefer a separately deployed helper, use the standalone ExcelDna.IntelliSense `.xll` from the project releases.

Native AOT add-ins should not reference this package directly. Deploy the standalone IntelliSense `.xll` alongside the Native AOT add-in instead.

More usage notes and examples are available at https://github.com/Excel-DNA/IntelliSense.
