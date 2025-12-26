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
        public static MatchEvaluator Replacer = (m) => new string('_', m.Value.Length);

        // TODO: What's the Unicode situation?
        public static string forbiddenNameCharacters = @"\ /\-:;!@\#\$%\^&\*\(\)\+=,<>\[\]{}|'\""";
        public static string functionNameRegex =
            "[" + forbiddenNameCharacters + "](?<functionName>[^" + forbiddenNameCharacters + "]*)$";
        public static string functionNameGroupName = "functionName";

        internal static bool TryGetFormulaInfo(
            FormulaEditWatcher.FormulaData formulaData,
            out string functionName,
            out int currentArgIndex,
            out string currentArgValue
        )
        {
            functionName = null;
            currentArgValue = null;
            currentArgIndex = -1;
            //Debug.Assert(formulaPrefix != null);
            if (formulaData.Text.Length < 2 || formulaData.CaretEndPosition < 2)
            {
                return false;
            }
            var originalFormulaPrefix = formulaData.Text;
            var formulaPrefix = formulaData.Text;
            var closedFunctionOffset = 0;

            // if (formulaPrefix.Length == formulaData.CaretEndPosition)
            // {
            //     closedFunctionOffset += 1;
            // }
            // if formula is closed open so that the following logic works
            if (formulaPrefix.Last() == ')' && formulaPrefix.Count(c => c == '(') == formulaPrefix.Count(c => c == ')'))
            {
                if (formulaData.Text.Length == formulaData.CaretEndPosition)
                    return false;
                formulaPrefix = formulaPrefix[..^1];
                closedFunctionOffset += 1;
            }

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
            formulaPrefix = Regex.Replace(formulaPrefix, "(\"[^\"]*\")", Replacer);

            // Remove sub-formulae
            formulaPrefix = Regex.Replace(formulaPrefix, "(\\([^\\(\\)]*\\))| ", Replacer);

            while (Regex.IsMatch(formulaPrefix, "\\([^\\(\\)]*\\)"))
            {
                formulaPrefix = Regex.Replace(formulaPrefix, "\\([^\\(\\)]*\\)", Replacer);
            }

            // Find the function name and the argument index
            int lastOpeningParenthesis = formulaPrefix.LastIndexOf(
                "(",
                formulaData.CaretEndPosition - 1,
                StringComparison.Ordinal
            );

            if (lastOpeningParenthesis > -1)
            {
                var match = Regex.Match(formulaPrefix.Substring(0, lastOpeningParenthesis), functionNameRegex);
                if (match.Success)
                {
                    functionName = match.Groups[functionNameGroupName].Value;

                    string argumentsPart = formulaPrefix.Substring(
                        lastOpeningParenthesis,
                        formulaData.CaretEndPosition - lastOpeningParenthesis
                    );
                    string originalArgumentsPart = argumentsPart;

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
                    argumentsPart = Regex.Replace(argumentsPart, "(\\{[^\\}]*\\})", Replacer);

                    currentArgIndex = argumentsPart.Count(c => c == ListSeparator);
                    if (currentArgIndex == 0)
                    {
                        currentArgValue = originalFormulaPrefix.Substring(lastOpeningParenthesis + 1);
                    }
                    else
                    {
                        var lastIndexOfComma = argumentsPart.LastIndexOf(ListSeparator);
                        var nextCommaAfterLast = originalFormulaPrefix.IndexOf(
                            ListSeparator,
                            lastIndexOfComma + lastOpeningParenthesis + 1
                        );
                        if (nextCommaAfterLast == -1)
                        {
                            nextCommaAfterLast = originalFormulaPrefix.Length - closedFunctionOffset;
                        }
                        currentArgValue = originalFormulaPrefix[
                            (lastIndexOfComma + lastOpeningParenthesis + 1)..nextCommaAfterLast
                        ];
                    }

                    return true;
                }
            }

            return false;
        }

        internal static string GetLineBeforeFunctionName(string formulaPrefix, string functionName)
        {
            // TODO: This should be cleaned up - just added something for testing...
            int lastFuncStart = formulaPrefix.LastIndexOf(functionName);
            if (lastFuncStart < 0)
                return "";

            int lastLineStart = formulaPrefix.LastIndexOf('\n', lastFuncStart) + 1;
            return formulaPrefix.Substring(lastLineStart, lastFuncStart - lastLineStart);
        }
    }
}
