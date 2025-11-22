using System.Collections.Generic;

namespace MySheets.Core.Models;

public class Worksheet {
    private readonly Dictionary<(int Row, int Col), Cell> _cells = new();
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
    }

    public int ActiveCellCount => _cells.Count;
}