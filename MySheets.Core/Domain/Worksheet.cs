using System.Globalization;
using System.Text.RegularExpressions;
using MySheets.Core.Calculation;
using MySheets.Core.Common;

namespace MySheets.Core.Domain;

public class Worksheet {
    private readonly Dictionary<(int Row, int Col), Cell> _cells = new();
    private readonly FormulaEvaluator _evaluator = new();
    
    public DependencyGraph DependencyGraph { get; private set; } = new();
    
    public event Action<int, int>? CellStateChanged;

    public IEnumerable<Cell> Cells => _cells.Values;

    public void Clear() {
        _cells.Clear();
        DependencyGraph = new DependencyGraph();
    }

    public Cell GetCell(int row, int col) {
        var key = (row, col);
        if (_cells.TryGetValue(key, out var cell)) {
            return cell;
        }
        var newCell = new Cell(row, col);
        _cells[key] = newCell;
        return newCell;
    }

    public bool IsFormulaValid(string formula) {
        if (string.IsNullOrEmpty(formula) || !formula.StartsWith("=")) return true;
        try {
            var result = _evaluator.Evaluate(formula, _ => 1.0);
            
            if (result is string s && s == "#ERROR!") return false;
            
            return true;
        } catch {
            return false;
        }
    }

    public void SetCell(int row, int col, string value) {
        DependencyGraph.ClearDependencies(row, col);
        var cell = GetCell(row, col);
        cell.Expression = value;

        UpdateCellInternal(row, col, cell);
        
        CellStateChanged?.Invoke(row, col);
        Recalculate(row, col);
        RecoverCyclicCells();
    }

    private void UpdateCellInternal(int row, int col, Cell cell) {
        if (cell.Expression.StartsWith("=")) {
            bool hasCycle = false;
            List<(int, int)> cyclePath = new();
            var referencesToAdd = new HashSet<(int, int)>();

            foreach (var reference in CellReferenceUtility.ExtractReferences(cell.Expression)) {
                referencesToAdd.Add(reference);
            }

            var rangeRegex = new Regex(@"([A-Za-z]+[0-9]+):([A-Za-z]+[0-9]+)");
            foreach (Match match in rangeRegex.Matches(cell.Expression)) {
                try {
                    foreach (var refPoint in CellReferenceUtility.ParseRange(match.Value)) {
                        referencesToAdd.Add(refPoint);
                    }
                } catch {  }
            }

            foreach (var reference in referencesToAdd) {
                if (!DependencyGraph.TryAddDependency(row, col, reference.Item1, reference.Item2, out var path)) {
                    hasCycle = true;
                    cyclePath = path;
                    break;
                }
            }

            if (hasCycle) {
                cell.Value = "#CYCLE!";
                foreach (var (cRow, cCol) in cyclePath) {
                    var cycleCell = GetCell(cRow, cCol);
                    cycleCell.Value = "#CYCLE!";
                    CellStateChanged?.Invoke(cRow, cCol);
                }
            } else {
                cell.Value = _evaluator.Evaluate(cell.Expression, GetCellValue);
            }
        } else {
            if (double.TryParse(cell.Expression, NumberStyles.Any, CultureInfo.InvariantCulture, out double numberResult)) {
                cell.Value = numberResult;
            } else {
                cell.Value = cell.Expression;
            }
        }
    }

    private void RecoverCyclicCells() {
        var cyclicCells = _cells.Values
            .Where(c => c.Value is string s && s == "#CYCLE!")
            .ToList();

        foreach (var cell in cyclicCells) {
            DependencyGraph.ClearDependencies(cell.Row, cell.Col);
            UpdateCellInternal(cell.Row, cell.Col, cell);
            
            if (cell.Value is not string || (string)cell.Value != "#CYCLE!") {
                CellStateChanged?.Invoke(cell.Row, cell.Col);
                Recalculate(cell.Row, cell.Col);
            }
        }
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

    private object GetCellValue(string cellReference) {
        try {
            var (row, col) = CellReferenceUtility.ParseReference(cellReference);
            var cell = GetCell(row, col);
            
            if (cell.Value is string s && s == "#CYCLE!") return "#CYCLE!";

            if (cell.Value is double d) return d;
            if (cell.Value is int i) return i;
            if (cell.Value == null) return 0.0;
            
            if (double.TryParse(cell.Value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double res)) {
                return res;
            }
            
            return 0.0; 
        } catch {
            throw new Exception("Invalid reference");
        }
    }

    public int ActiveCellCount => _cells.Count;
}