using System;
using Avalonia;
using Avalonia.Layout; 
using Avalonia.Media;
using CommunityToolkit.Mvvm.ComponentModel;
using MySheets.Core.Domain;
using MySheets.Core.Common;

namespace MySheets.UI.ViewModels.SheetEditor;

public class CellViewModel : ObservableObject {
    private readonly Cell _model;
    private readonly Worksheet _sheet;
    private readonly SheetViewModel _parentVm;

    public CellViewModel(Cell model, Worksheet sheet, ColumnViewModel column, SheetViewModel parentVm) {
        _model = model;
        _sheet = sheet;
        _parentVm = parentVm;
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
                var oldValue = _model.Expression;
                var newValue = value;
                
                // Record action
                _parentVm.History.Execute(new CellEditAction(
                    oldValue, 
                    newValue, 
                    val => {
                        _sheet.SetCell(_model.Row, _model.Col, val);
                        OnPropertyChanged(nameof(Expression));
                    }));

                // Apply immediately
                _sheet.SetCell(_model.Row, _model.Col, newValue);
                OnPropertyChanged();
            }
        }
    }
    
    public double FontSize {
        get => _model.FontSize;
        set {
            if (Math.Abs(_model.FontSize - value) > 0.01) {
                CaptureStyleChange(_model.FontSize, value, v => {
                    _model.FontSize = v;
                    OnPropertyChanged(nameof(FontSize));
                });
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
                CaptureStyleChange(_model.IsBold, value, v => {
                    _model.IsBold = v;
                    OnPropertyChanged(nameof(IsBold));
                    OnPropertyChanged(nameof(FontWeight));
                });
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
                CaptureStyleChange(_model.IsItalic, value, v => {
                    _model.IsItalic = v;
                    OnPropertyChanged(nameof(IsItalic));
                    OnPropertyChanged(nameof(FontStyle));
                });
                _model.IsItalic = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(FontStyle)); 
            }
        }
    }

    public FontStyle FontStyle => IsItalic ? FontStyle.Italic : FontStyle.Normal;
    
    public IBrush Foreground => Brush.Parse(_model.TextColor);
    
    public IBrush Background => Brush.Parse(_model.BackgroundColor);

    public Thickness BorderThickness {
        get {
            if (_model.BorderThickness == "0,0,1,1") {
                return new Thickness(0);
            }
            return Thickness.Parse(_model.BorderThickness);
        }
    }
    
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
                CaptureStyleChange(_model.TextAlignment, newVal, v => {
                    _model.TextAlignment = v;
                    OnPropertyChanged(nameof(CellAlignment));
                });
                _model.TextAlignment = newVal;
                OnPropertyChanged();
            }
        }
    }
    
    public void SetTextColor(string colorHex) {
        if (_model.TextColor != colorHex) {
            CaptureStyleChange(_model.TextColor, colorHex, v => {
                _model.TextColor = v;
                OnPropertyChanged(nameof(Foreground));
            });
            _model.TextColor = colorHex;
            OnPropertyChanged(nameof(Foreground));
        }
    }

    public void SetBackgroundColor(string colorHex) {
        if (_model.BackgroundColor != colorHex) {
            CaptureStyleChange(_model.BackgroundColor, colorHex, v => {
                _model.BackgroundColor = v;
                OnPropertyChanged(nameof(Background));
            });
            _model.BackgroundColor = colorHex;
            OnPropertyChanged(nameof(Background));
        }
    }

    public void SetBorder(string thickness) {
        if (_model.BorderThickness != thickness) {
            CaptureStyleChange(_model.BorderThickness, thickness, v => {
                _model.BorderThickness = v;
                OnPropertyChanged(nameof(BorderThickness));
            });
            _model.BorderThickness = thickness;
            OnPropertyChanged(nameof(BorderThickness));
        }
    }

    private void CaptureStyleChange<T>(T oldVal, T newVal, Action<T> apply) {
        _parentVm.History.Execute(new CellStyleAction<T>(oldVal, newVal, apply));
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