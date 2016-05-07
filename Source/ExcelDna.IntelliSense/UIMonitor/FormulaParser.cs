using System.Linq;
using System.Text.RegularExpressions;

namespace ExcelDna.IntelliSense
{
    static class FormulaParser
    {
        // TODO: This needs a proper implementation, considering subformulae
        internal static bool TryGetFormulaInfo(string formulaPrefix, out string functionName, out int currentArgIndex)
        {
            var match = Regex.Match(formulaPrefix, @"^=(?<functionName>(\w|.)*)\(");
            if (match.Success)
            {
                functionName = match.Groups["functionName"].Value;
                currentArgIndex = formulaPrefix.Count(c => c == ',');
                return true;
            }
            functionName = null;
            currentArgIndex = -1;
            return false;
        }
    }
}
