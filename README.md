Excel-DNA IntelliSense
======================
Excel-DNA - see http://excel-dna.net - is an independent project to integrate .NET with Excel.
With Excel-DNA you can make native (.xll) add-ins for Excel using C#, Visual Basic.NET or F#, providing high-performance user-defined functions (UDFs), custom ribbon interfaces and more.

This project adds in-sheet IntelliSense for Excel UDFs as part of an Excel-DNA add-in.

Overview
--------
Excel has no known support for user-defined functions to display as part of the on-sheet intellisense. We use the UI Automation support of Windows and Excel, to keep track of relevant changes of the Excel interface, and overlay IntelliSense information when appropriate.

Current status
--------------
The project is under activate development, and at a very early preview stage.
As a proof of concept we have the following.

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

Future direction
----------------

The first step is to stabilize the current implementation, and resolve how the feature will be distributed. Since the IntelliSense extension can onyl be provided by one add-in loaded into Excel, different add-ins providing the service will have to cooperate.

Once a basic implementation is working, there is scope for quite a lot of enhancement. For example, we could add support for:

  * enum lists and other parameter selection and validation
  * links to forms or hyperlinks to help
  * enhanced argument selection controls, like a date selector

The intention is also that the Excel-DNA IntelliSense helper could be used both to add IntelliSense help for UDFs defined in Excel-DNA based add-ins, and for VBA-based functions. We'd need to use some kind of definition file or registration function to allow the VBA function to register the descriptions.

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
  
  17 April 2015
  
