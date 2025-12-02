namespace MySheets.UI.ViewModels;

using CommunityToolkit.Mvvm.ComponentModel;
using MySheets.Core.Models;

public class CellViewModel : ObservableObject {
    private readonly Cell _model;
    private readonly Worksheet _sheet;

    public CellViewModel(Cell model, Worksheet sheet, ColumnViewModel column) {
        _model = model;
        _sheet = sheet;
        Column = column;
        
        _sheet.CellStateChanged += OnCellStateChanged;
    }

    private void OnCellStateChanged(int row, int col) {
        if (row == _model.Row && col == _model.Col) {
            OnPropertyChanged(nameof(Value));
        }
    }

    public ColumnViewModel Column { get; }

    public int Row => _model.Row;

    public int Col => _model.Col;

    public string Expression {
        get => _model.Expression;
        set {
            if (_model.Expression != value) {
                _sheet.SetCell(_model.Row, _model.Col, value);
                OnPropertyChanged();
            }
        }
    }

    public object Value => _model.Value;
}