using System.Text.RegularExpressions;

namespace MySheets.Core.Utilities;

public static class CellReferenceUtility {
    private static readonly Regex ReferenceRegex = new(@"([A-Z]+)([0-9]+)");

    public static (int Row, int Col) ParseReference(string reference) {
        var match = ReferenceRegex.Match(reference);
        if (!match.Success) {
            throw new ArgumentException($"Invalid cell reference: {reference}");
        }

        var colStr = match.Groups[1].Value;
        var rowStr = match.Groups[2].Value;

        var colIndex = GetColumnIndex(colStr);
        var rowIndex = int.Parse(rowStr) - 1;

        return (rowIndex, colIndex);
    }

    public static IEnumerable<(int Row, int Col)> ExtractReferences(string formula) {
        var matches = ReferenceRegex.Matches(formula);
        foreach (Match match in matches) {
            yield return ParseReference(match.Value);
        }
    }

    private static int GetColumnIndex(string columnName) {
        int sum = 0;
        foreach (var c in columnName) {
            sum *= 26;
            sum += (c - 'A' + 1);
        }
        return sum - 1;
    }
}