Excel-DNA IntelliSense
======================
Excel-DNA - see http://excel-dna.net - is an independent project to integrate .NET with Excel.
With Excel-DNA you can make native (.xll) add-ins for Excel using C#, Visual Basic.NET or F#, providing high-performance user-defined functions (UDFs), custom ribbon interfaces and more.

This project adds in-sheet IntelliSense for Excel UDFs, either through an independently deployed add-in or as part of an Excel-DNA add-in.

Overview
--------
Excel has no known support for user-defined functions to display as part of the on-sheet intellisense. We use the UI Automation support of Windows and Excel, to keep track of relevant changes of the Excel interface, and overlay IntelliSense information when appropriate.

Current status
--------------
The project is under activate development, and ready for intial testing.

For an Excel-DNA function defined like this:
```C#
[ExcelFunction(Description = "A useful test function that adds two numbers, and returns the sum.")]
public static double AddThem(
	[ExcelArgument(Name = "Augend", Description = "is the first number, to which will be added")] 
	double v1,
	[ExcelArgument(Name = "Addend", Description = "is the second number that will be added")]     
	double v2)
{
	return v1 + v2;
}
```
we get both the function description

![Function Description](https://raw.github.com/Excel-DNA/IntelliSense/master/Screenshots/FunctionDescription.PNG)

and when selecting the function, we get argument help

![Argument Help](https://raw.github.com/Excel-DNA/IntelliSense/master/Screenshots/ArgumentHelp.PNG)


User-defined functions written in VBA (either in an add-in or regular Workbook) can also provide IntelliSense descriptions, either by embedding descriptions in the Workbook, or in an external file.

The first configuration being tested now, is where the IntelliSense display server is loaded as a separate add-in.

Getting Started
---------------

For existing Excel-DNA add-ins (v0.32 or later):
  * Download and load the latest ExcelDna.IntelliSense.xll or ExcelDna.IntelliSense64.xll from the [Releases](https://github.com/Excel-DNA/IntelliSense/releases) page.
  * IntelliSense should work automatically for functions that have descriptions in [ExcelFunction] and [ExcelArgument] attributes.

For VBA workbooks or add-ins:
  * Download and load the latest ExcelDna.IntelliSense.xll or ExcelDna.IntelliSense64.xll from the [Releases](https://github.com/Excel-DNA/IntelliSense/releases) page.
  * Either add a sheet with the IntelliSense function descriptions, or a separate xml file.

See the [Getting Started](https://github.com/Excel-DNA/IntelliSense/wiki/Getting-Started) page for more detail.

Future direction
----------------

Once a basic implementation is working, there is scope for quite a lot of enhancement. For example, we could add support for:

  * enum lists and other parameter selection and validation
  * links to forms or hyperlinks to help
  * enhanced argument selection controls, like a date selector

Support and participation
-------------------------
"We accept pull requests" ;-) 
Any help or feedback is greatly appreciated.

Please log bugs and feature suggestions on the GitHub 'Issues' page.

For general comments or discussion, use the Excel-DNA forum at https://groups.google.com/forum/#!forum/exceldna .

License
-------
This project is published under the standard MIT license.


  Govert van Drimmelen
  
  govert@icon.co.za
  
  18 June 2016
  
