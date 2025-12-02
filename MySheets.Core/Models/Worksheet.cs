using System.Globalization;
using MySheets.Core.Services;
using MySheets.Core.Utilities;

namespace MySheets.Core.Models;

public class Worksheet {
    private readonly Dictionary<(int Row, int Col), Cell> _cells = new();
    private readonly FormulaEvaluator _evaluator = new();
    public DependencyGraph DependencyGraph { get; } = new();

    public Cell GetCell(int row, int col) {
        var key = (row, col);
        if (_cells.TryGetValue(key, out var cell)) {
            return cell;
        }
        var newCell = new Cell(row, col);
        _cells[key] = newCell;
        return newCell;
    }

    public void SetCell(int row, int col, string value) {
        DependencyGraph.ClearDependencies(row, col);
        var cell = GetCell(row, col);
        cell.Expression = value;

        if (value.StartsWith("=")) {
            foreach (var reference in CellReferenceUtility.ExtractReferences(value)) {
                DependencyGraph.AddDependency(row, col, reference.Row, reference.Col);
            }
            cell.Value = _evaluator.Evaluate(value);
        } else {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numberResult)) {
                cell.Value = numberResult;
            } else {
                cell.Value = value;
            }
        }
    }

    public int ActiveCellCount => _cells.Count;
}