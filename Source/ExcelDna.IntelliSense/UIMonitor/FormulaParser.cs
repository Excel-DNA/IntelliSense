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
            formulaPrefix = Regex.Replace(formulaPrefix, "(\"[^\"]*\")|(\\([^\\(\\)]*\\))| ", string.Empty);

            while (Regex.IsMatch(formulaPrefix, "\\([^\\(\\)]*\\)"))
            {
                formulaPrefix = Regex.Replace(formulaPrefix, "\\([^\\(\\)]*\\)", string.Empty);
            }

            int lastOpeningParenthesis = formulaPrefix.LastIndexOf("(", formulaPrefix.Length - 1, StringComparison.Ordinal);

            if (lastOpeningParenthesis > -1)
            {
                var match = Regex.Match(formulaPrefix.Substring(0, lastOpeningParenthesis), functionNameRegex);
                if (match.Success)
                {
                    functionName = match.Groups[functionNameGroupName].Value;
                    currentArgIndex = formulaPrefix.Substring(lastOpeningParenthesis, formulaPrefix.Length - lastOpeningParenthesis).Count(c => c == ListSeparator);
                    return true;
                }
            }

            functionName = null;
            currentArgIndex = -1;
            return false;
        }
    }
}
