using System;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelDna.IntelliSense
{
    static class FormulaParser
    {
        // Set from IntelliSenseDisplay.Initialize
        public static char ListSeparator = ',';
        // TODO: What's the Unicode situation?
        public static string forbiddenNameCharacters = @"\ /\-:;!@\#\$%\^&\*\(\)\+=,<>\[\]{}|'\""";
        public static string functionNameRegex = "[" + forbiddenNameCharacters + "](?<functionName>[^" + forbiddenNameCharacters + "]*)$";
        public static string functionNameGroupName = "functionName";

        internal static bool TryGetFormulaInfo(string formulaPrefix, out string functionName, out int currentArgIndex)
        {
            Debug.Assert(formulaPrefix != null);

            // Hide the strings, in order to ignore the commas, parenthesis and curly brackets which might be inside.
            // Ex: =SomeFunction("a", "b,c(d,{e,f, 
            // In the example above, the function is SomeFunction, and the index is 2

            // If editing a string, then close the double-quotes
            if (formulaPrefix.Count(c => c == '\"') % 2 != 0)
            {
                formulaPrefix = string.Concat(formulaPrefix, '\"');
            }
            
            // Remove the strings.
            // Note: in Excel, in order to put a double-quotes in a string, one has to double the double-quotes.
            // For instance, "a""b" for a"b. Since we only want to hide the strings in order to count the commas, 
            // the regex below still applies.
            formulaPrefix = Regex.Replace(formulaPrefix, "(\"[^\"]*\")", string.Empty);

            // Remove sub-formulae
            formulaPrefix = Regex.Replace(formulaPrefix, "(\\([^\\(\\)]*\\))| ", string.Empty);

            while (Regex.IsMatch(formulaPrefix, "\\([^\\(\\)]*\\)"))
            {
                formulaPrefix = Regex.Replace(formulaPrefix, "\\([^\\(\\)]*\\)", string.Empty);
            }

            // Find the function name and the argument index
            int lastOpeningParenthesis = formulaPrefix.LastIndexOf("(", formulaPrefix.Length - 1, StringComparison.Ordinal);

            if (lastOpeningParenthesis > -1)
            {
                var match = Regex.Match(formulaPrefix.Substring(0, lastOpeningParenthesis), functionNameRegex);
                if (match.Success)
                {
                    functionName = match.Groups[functionNameGroupName].Value;

                    string argumentsPart = formulaPrefix.Substring(lastOpeningParenthesis, formulaPrefix.Length - lastOpeningParenthesis);

                    // Hide array formulae
                    // Ex: =SomeFunction("a", {"a", "b", "c"
                    // In the example above the index is 2

                    // If editing an array, then close the curly bracket.
                    // Since we already removed the strings, there won't be any curly brackets within a string.
                    if (argumentsPart.Count(c => c == '{') > argumentsPart.Count(c => c == '}'))
                    {
                        argumentsPart = string.Concat(argumentsPart, '}');
                    }

                    // Remove the arrays.
                    argumentsPart = Regex.Replace(argumentsPart, "(\\{[^\\}]*\\})", string.Empty);

                    currentArgIndex = argumentsPart.Count(c => c == ListSeparator);
                    return true;
                }
            }

            functionName = null;
            currentArgIndex = -1;
            return false;
        }
    }
}
