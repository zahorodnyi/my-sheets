namespace MySheets.UI.Views;

using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using MySheets.UI.ViewModels;

public partial class MainWindow : Window {
    private bool _isResizingColumn;
    private bool _isResizingRow;
    private Point _lastMousePosition;
    private ColumnViewModel? _targetColumn;
    private RowViewModel? _targetRow;

    public MainWindow() {
        InitializeComponent();

        MainScrollViewer.ScrollChanged += (s, e) => {
            ColHeaderScrollViewer.Offset = new Vector(MainScrollViewer.Offset.X, 0);
            RowHeaderScrollViewer.Offset = new Vector(0, MainScrollViewer.Offset.Y);
        };
    }

    private void OnColumnHeaderPressed(object? sender, PointerPressedEventArgs e) {
        if (sender is Border border && border.Tag is ColumnViewModel column) {
            _isResizingColumn = true;
            _targetColumn = column;
            _lastMousePosition = e.GetPosition(this);
            e.Pointer.Capture(border);
        }
    }

    private void OnRowHeaderPressed(object? sender, PointerPressedEventArgs e) {
        if (sender is Border border && border.Tag is RowViewModel row) {
            _isResizingRow = true;
            _targetRow = row;
            _lastMousePosition = e.GetPosition(this);
            e.Pointer.Capture(border);
        }
    }

    private void OnPointerMoved(object? sender, PointerEventArgs e) {
        if (!_isResizingColumn && !_isResizingRow) return;

        var currentPosition = e.GetPosition(this);
        var delta = _isResizingColumn 
            ? currentPosition.X - _lastMousePosition.X 
            : currentPosition.Y - _lastMousePosition.Y;

        if (_isResizingColumn && _targetColumn != null) {
            _targetColumn.Width = Math.Max(30, _targetColumn.Width + delta);
        } else if (_isResizingRow && _targetRow != null) {
            _targetRow.Height = Math.Max(20, _targetRow.Height + delta);
        }

        _lastMousePosition = currentPosition;
    }

    private void OnPointerReleased(object? sender, PointerReleasedEventArgs e) {
        if (_isResizingColumn) {
            _isResizingColumn = false;
            _targetColumn = null;
        }

        if (_isResizingRow) {
            _isResizingRow = false;
            _targetRow = null;
        }
    }
}