using System.Globalization;
using MySheets.Core.Services;
using MySheets.Core.Utilities;

namespace MySheets.Core.Models;

public class Worksheet {
    private readonly Dictionary<(int Row, int Col), Cell> _cells = new();
    private readonly FormulaEvaluator _evaluator = new();
    
    public DependencyGraph DependencyGraph { get; } = new();
    
    public event Action<int, int>? CellStateChanged;

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
            cell.Value = _evaluator.Evaluate(value, GetCellValue);
        } else {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numberResult)) {
                cell.Value = numberResult;
            } else {
                cell.Value = value;
            }
        }
        
        CellStateChanged?.Invoke(row, col);
        Recalculate(row, col);
    }

    private void Recalculate(int row, int col) {
        foreach (var dependent in DependencyGraph.GetDependents(row, col)) {
            var cell = GetCell(dependent.Item1, dependent.Item2);
            
            if (cell.Type == CellType.Formula) {
                cell.Value = _evaluator.Evaluate(cell.Expression, GetCellValue);
                CellStateChanged?.Invoke(dependent.Item1, dependent.Item2);
                Recalculate(dependent.Item1, dependent.Item2);
            }
        }
    }

    private double GetCellValue(string cellReference) {
        try {
            var (row, col) = CellReferenceUtility.ParseReference(cellReference);
            var cell = GetCell(row, col);
            
            if (cell.Value is double d) return d;
            if (cell.Value is int i) return i;
            if (cell.Value == null) return 0;
            
            if (double.TryParse(cell.Value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double res)) {
                return res;
            }
            
            return 0; 
        } catch {
            throw new Exception("Invalid reference");
        }
    }

    public int ActiveCellCount => _cells.Count;
}