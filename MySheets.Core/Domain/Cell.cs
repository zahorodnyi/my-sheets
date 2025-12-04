using System.Globalization;

namespace MySheets.Core.Domain;

public enum CellType {
    Text,
    Number,
    Formula
}

public class Cell {
    public int Row { get; }
    public int Col { get; }

    private string _expression = string.Empty;

    public string Expression {
        get => _expression;
        set {
            _expression = value;
            DetermineType();
        }
    }
    public object? Value { get; set; }
    public double FontSize { get; set; } = 12.0; 
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public string TextColor { get; set; } = "#000000"; 
    public string BackgroundColor { get; set; } = "Transparent";
    public string BorderThickness { get; set; } = "0,0,1,1";

    public CellType Type { get; private set; }

    public Cell(int row, int column) {
        if (row < 0 || column < 0)
            throw new ArgumentOutOfRangeException(nameof(row), "Coordinates cannot be negative.");

        Row = row;
        Col = column;
        DetermineType();
    }

    private void DetermineType() {
        if (string.IsNullOrEmpty(Expression)) {
            Type = CellType.Text;
            return;
        }

        if (Expression.StartsWith('=')) {
            Type = CellType.Formula;
        } else if (double.TryParse(Expression, NumberStyles.Any, CultureInfo.InvariantCulture, out double numberResult)) {
            Type = CellType.Number;
            Value = numberResult; 
        } else {
            Type = CellType.Text;
            Value = Expression;
        }
    }
}