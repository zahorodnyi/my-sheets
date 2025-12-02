namespace MySheets.UI.Views;

using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using MySheets.UI.ViewModels;

public partial class MainWindow : Window {
    private bool _isResizingColumn;
    private bool _isResizingRow;
    private bool _isSelecting;
    
    private Point _lastMousePosition;
    private ColumnViewModel? _targetColumn;
    private RowViewModel? _targetRow;
    
    private Control? _capturedControl; 
    
    private int _lastSelectedRowIndex = -1;
    private int _lastSelectedColIndex = -1;

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
            _capturedControl = border;
            e.Pointer.Capture(border);
            e.Handled = true;
        }
    }

    private void OnRowHeaderPressed(object? sender, PointerPressedEventArgs e) {
        if (sender is Border border && border.Tag is RowViewModel row) {
            _isResizingRow = true;
            _targetRow = row;
            _lastMousePosition = e.GetPosition(this);
            _capturedControl = border;
            e.Pointer.Capture(border);
            e.Handled = true;
        }
    }

    private void OnMainPointerPressed(object? sender, PointerPressedEventArgs e) {
        var properties = e.GetCurrentPoint(this).Properties;
        if (!properties.IsLeftButtonPressed) return;
        if (_isResizingColumn || _isResizingRow) return;

        if (DataContext is MainWindowViewModel vm && sender is Control panel) {
            var mousePoint = e.GetPosition(MainScrollViewer);
            
            double absX = mousePoint.X + MainScrollViewer.Offset.X;
            double absY = mousePoint.Y + MainScrollViewer.Offset.Y;
            
            var (r, c) = GetRowColumnAt(absX, absY, vm);
            
            if (r != -1 && c != -1) {
                if (r == _lastSelectedRowIndex && c == _lastSelectedColIndex) {
                    EnterEditMode(panel, e.GetPosition(panel));
                    return;
                }

                _isSelecting = true;
                _lastSelectedRowIndex = r;
                _lastSelectedColIndex = c;
                
                vm.StartSelection(r, c);
                
                _capturedControl = panel;
                e.Pointer.Capture(panel);
            }
        }
    }
    
    private void EnterEditMode(Control parentControl, Point point) {
        var hit = parentControl.InputHitTest(point);
        
        if (hit is Visual visual) {
            var cellBorder = (visual as Border) ?? visual.FindAncestorOfType<Border>();
            if (cellBorder != null) {
                var textBox = cellBorder.FindDescendantOfType<TextBox>();
                if (textBox != null) {
                    textBox.IsHitTestVisible = true;
                    textBox.Focus();
                }
            }
        }
    }

    private void OnCellEditorLostFocus(object? sender, RoutedEventArgs e) {
        if (sender is TextBox textBox) {
            textBox.IsHitTestVisible = false;
        }
    }

    private void OnPointerMoved(object? sender, PointerEventArgs e) {
        if (_isResizingColumn && _targetColumn != null) {
            var currentPosition = e.GetPosition(this);
            var delta = currentPosition.X - _lastMousePosition.X;
            _targetColumn.Width = Math.Max(30, _targetColumn.Width + delta);
            _lastMousePosition = currentPosition;
        } else if (_isResizingRow && _targetRow != null) {
            var currentPosition = e.GetPosition(this);
            var delta = currentPosition.Y - _lastMousePosition.Y;
            _targetRow.Height = Math.Max(20, _targetRow.Height + delta);
            _lastMousePosition = currentPosition;
        } 
        else if (_isSelecting && DataContext is MainWindowViewModel vm) {
            var mousePoint = e.GetPosition(MainScrollViewer);
            
            double absX = mousePoint.X + MainScrollViewer.Offset.X;
            double absY = mousePoint.Y + MainScrollViewer.Offset.Y;

            var (r, c) = GetRowColumnAt(absX, absY, vm);
            if (r != -1 && c != -1) {
                vm.UpdateSelection(r, c);
                
                _lastSelectedRowIndex = r;
                _lastSelectedColIndex = c;
            }
        }
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

        if (_isSelecting) {
            _isSelecting = false;
        }

        if (_capturedControl != null) {
            e.Pointer.Capture(null);
            _capturedControl = null;
        }
    }

    private (int row, int col) GetRowColumnAt(double x, double y, MainWindowViewModel vm) {
        int colIndex = -1;
        double currentX = 0;
        
        for (int i = 0; i < vm.ColumnHeaders.Count; i++) {
            double w = vm.ColumnHeaders[i].Width;
            if (x >= currentX && x < currentX + w) {
                colIndex = i;
                break;
            }
            currentX += w;
        }

        int rowIndex = -1;
        double currentY = 0;

        for (int i = 0; i < vm.Rows.Count; i++) {
            double h = vm.Rows[i].Height;
            if (y >= currentY && y < currentY + h) {
                rowIndex = i;
                break;
            }
            currentY += h;
        }
        
        if (colIndex == -1 && x >= currentX && vm.ColumnHeaders.Count > 0) 
            colIndex = vm.ColumnHeaders.Count - 1;
        
        if (rowIndex == -1 && y >= currentY && vm.Rows.Count > 0) 
            rowIndex = vm.Rows.Count - 1;

        return (rowIndex, colIndex);
    }
}