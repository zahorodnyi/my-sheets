using System.Text.RegularExpressions;

namespace MySheets.Core.Common;

public static class CellReferenceUtility {
    private static readonly Regex ReferenceRegex = new(@"([A-Za-z]+)([0-9]+)");

    public static (int Row, int Col) ParseReference(string reference) {
        var match = ReferenceRegex.Match(reference);
        if (!match.Success) {
            throw new ArgumentException($"Invalid cell reference: {reference}");
        }

        var colStr = match.Groups[1].Value.ToUpper();
        var rowStr = match.Groups[2].Value;

        var colIndex = GetColumnIndex(colStr);
        var rowIndex = int.Parse(rowStr) - 1;

        return (rowIndex, colIndex);
    }

    public static IEnumerable<(int Row, int Col)> ExtractReferences(string formula) {
        var matches = ReferenceRegex.Matches(formula);
        foreach (Match match in matches) {
            if (match.Index > 0 && formula[match.Index - 1] == ':') continue; 
            if (match.Index + match.Length < formula.Length && formula[match.Index + match.Length] == ':') continue;
            
            yield return ParseReference(match.Value);
        }
    }

    public static IEnumerable<(int Row, int Col)> ParseRange(string rangeExpression) {
        var parts = rangeExpression.Split(':');
        if (parts.Length != 2) throw new ArgumentException("Invalid range format");

        var start = ParseReference(parts[0]);
        var end = ParseReference(parts[1]);

        int r1 = Math.Min(start.Row, end.Row);
        int r2 = Math.Max(start.Row, end.Row);
        int c1 = Math.Min(start.Col, end.Col);
        int c2 = Math.Max(start.Col, end.Col);

        for (int r = r1; r <= r2; r++) {
            for (int c = c1; c <= c2; c++) {
                yield return (r, c);
            }
        }
    }

    public static string GetColumnName(int index) {
        int dividend = index + 1;
        string columnName = string.Empty;
        while (dividend > 0) {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            dividend = (int)((dividend - modulo) / 26);
        }
        return columnName;
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