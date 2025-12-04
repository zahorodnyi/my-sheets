namespace MySheets.UI.ViewModels;

using System;
using CommunityToolkit.Mvvm.ComponentModel;
using MySheets.Core.Models;
using Avalonia.Media;

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
    
    public double FontSize {
        get => _model.FontSize;
        set {
            if (Math.Abs(_model.FontSize - value) > 0.01) {
                _model.FontSize = value;
                OnPropertyChanged();
            }
        }
    }
    
    public object Value {
        get {
            var val = _model.Value;
            if (val is string str && str.StartsWith("'")) {
                return str.Substring(1);
            }
            return val;
        }
    }
    
    public bool IsBold {
        get => _model.IsBold;
        set {
            if (_model.IsBold != value) {
                _model.IsBold = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(FontWeight)); 
            }
        }
    }

    public FontWeight FontWeight => IsBold ? FontWeight.Bold : FontWeight.Normal;

    public bool IsItalic {
        get => _model.IsItalic;
        set {
            if (_model.IsItalic != value) {
                _model.IsItalic = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(FontStyle)); 
            }
        }
    }

    public FontStyle FontStyle => IsItalic ? FontStyle.Italic : FontStyle.Normal;
    
    public void Refresh() {
        OnPropertyChanged(nameof(Value));
        OnPropertyChanged(nameof(Expression));
        OnPropertyChanged(nameof(FontSize));
        OnPropertyChanged(nameof(FontWeight));
        OnPropertyChanged(nameof(FontStyle));
    }
}