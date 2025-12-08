using System;
using Avalonia;
using Avalonia.Layout; 
using Avalonia.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using MySheets.Core.Domain;

namespace MySheets.UI.ViewModels.SheetEditor;

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
    
    public IBrush Foreground => Brush.Parse(_model.TextColor);
    
    public IBrush Background => Brush.Parse(_model.BackgroundColor);
    
    public Thickness BorderThickness => Thickness.Parse(_model.BorderThickness);
    
    public HorizontalAlignment CellAlignment {
        get {
            return _model.TextAlignment switch {
                "Center" => HorizontalAlignment.Center,
                "Right" => HorizontalAlignment.Right,
                _ => HorizontalAlignment.Left
            };
        }
        set {
            string newVal = value.ToString();
            if (_model.TextAlignment != newVal) {
                _model.TextAlignment = newVal;
                OnPropertyChanged();
            }
        }
    }
    
    public void SetTextColor(string colorHex) {
        if (_model.TextColor != colorHex) {
            _model.TextColor = colorHex;
            OnPropertyChanged(nameof(Foreground));
        }
    }

    public void SetBackgroundColor(string colorHex) {
        if (_model.BackgroundColor != colorHex) {
            _model.BackgroundColor = colorHex;
            OnPropertyChanged(nameof(Background));
        }
    }

    public void SetBorder(string thickness) {
        if (_model.BorderThickness != thickness) {
            _model.BorderThickness = thickness;
            OnPropertyChanged(nameof(BorderThickness));
        }
    }
    
    public void Refresh() {
        OnPropertyChanged(nameof(Value));
        OnPropertyChanged(nameof(Expression));
        OnPropertyChanged(nameof(FontSize));
        OnPropertyChanged(nameof(FontWeight));
        OnPropertyChanged(nameof(FontStyle));
        OnPropertyChanged(nameof(Foreground));
        OnPropertyChanged(nameof(Background));
        OnPropertyChanged(nameof(BorderThickness));
        OnPropertyChanged(nameof(CellAlignment)); 
    }
}