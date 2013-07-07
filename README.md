Excel-DNA IntelliSense
======================
Excel-DNA - see http://excel-dna.net and http://exceldna.codeplex.com - is an independent project to integrate .NET with Excel.
With Excel-DNA you can make native (.xll) add-ins for Excel using C#, Visual Basic.NET or F#, providing high-performance user-defined functions (UDFs), custom ribbon interfaces and more.

This project adds in-sheet IntelliSense for Excel UDFs as part of an Excel-DNA add-in.

Overview
--------
Excel has no known support for user-defined functions to display as part of the on-sheet intellisense. We use the UI Automation support of Windows and Excel, to keep track of relevant changes of the Excel interface, and overlay IntelliSense information when appropriate.

Current status
--------------
At the moment we have an early preview that serves as a proof-of-concept that 'works on my machine', under Excel 2013 on Windows 8. I have tested on a 64-bit Excel 2010 on Windows Server 2008, with only partial success - function descriptions work, but I was not able to get the UI Automation TextPattern to work at all on that platform - this is needed for the formula help to be displayed.

Even on my machine, the window management is still problematic - the IntelliSense form gets detached from the Excel window, and disappears. At most we can show a promising direction...

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
https://raw.github.com/Excel-DNA/IntelliSense/master/Screenshots/FunctionDescription.png

and argument help
https://raw.github.com/Excel-DNA/IntelliSense/master/Screenshots/ArgumentHelp.png

Future direction
----------------
Moving beyond the proof-of-concept we need to check different Excel and Windows versions.

The intention is that the Excel-DNA IntelliSense helper could be used both to add IntelliSense help for UDF defined in Excel-DNA based add-ins, and for VBA-based functions. We'd need to use some kind of definition file or registration function to allow the VBA function to register the descriptions.

Support and participation
-------------------------
"We accept pull requests" ;-)
Please log bugs and feature suggestions on the GitHub 'Issues' page.
For general comments or discussion, use the Excel-DNA forum at https://groups.google.com/forum/#!forum/exceldna .

License
-------
This project is published under the standard MIT license.


  Govert van Drimmelen
  govert@icon.co.za
  8 July 2013
  
