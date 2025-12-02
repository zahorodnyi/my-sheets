using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using MySheets.UI.ViewModels;
using Avalonia.Platform.Storage;
using MySheets.UI.ViewModels;

namespace MySheets.UI.Views;

public partial class MainWindow : Window {
    private bool _isResizingColumn;
    private bool _isResizingRow;
    private bool _isSelecting;
    
    private Point _lastMousePosition;
    private ColumnViewModel? _targetColumn;
    private RowViewModel? _targetRow;
    private Control? _capturedControl; 
    
    private int _anchorRowIndex = -1;
    private int _anchorColIndex = -1;

    private TextBox? _activeFormulaEditor; 
    
    private bool _isLastActionRefSelect = false;
    private int _lastInsertedRefIndex = -1;
    private int _lastInsertedRefLength = 0;

    public MainWindow() {
        InitializeComponent();

        MainScrollViewer.ScrollChanged += (s, e) => {
            ColHeaderScrollViewer.Offset = new Vector(MainScrollViewer.Offset.X, 0);
            RowHeaderScrollViewer.Offset = new Vector(0, MainScrollViewer.Offset.Y);
        };
    }

    protected override void OnTextInput(TextInputEventArgs e) {
        base.OnTextInput(e);
        
        _isLastActionRefSelect = false;
        if (DataContext is MainWindowViewModel vm) vm.HideRefSelection();

        if (_activeFormulaEditor != null && _activeFormulaEditor.IsFocused) return; 

        if (!string.IsNullOrEmpty(e.Text) && 
            _anchorRowIndex != -1 && 
            _anchorColIndex != -1 && 
            DataContext is MainWindowViewModel viewModel) {
            
            var rect = GetCellRect(_anchorRowIndex, _anchorColIndex, viewModel);
            var centerPoint = new Point(rect.X + rect.Width / 2, rect.Y + rect.Height / 2);
            
            var textBox = EnterEditMode(CellPanel, centerPoint);
            
            if (textBox != null) {
                textBox.Text = e.Text;
                textBox.CaretIndex = e.Text.Length;
            }
        }
    }
    
    protected override void OnKeyDown(KeyEventArgs e) {
        base.OnKeyDown(e);
        if (e.Key == Key.Enter) {
            if (_activeFormulaEditor != null) {
                _activeFormulaEditor.IsHitTestVisible = false;
                _activeFormulaEditor.Opacity = 0;
                _activeFormulaEditor = null;
                
                if (DataContext is MainWindowViewModel vm) vm.HideRefSelection();
                
                CellPanel.Focus();
            }
            return;
        }

        if (_activeFormulaEditor != null && _activeFormulaEditor.IsFocused) {
            if (e.Key != Key.None) {
                 _isLastActionRefSelect = false;
                 if (DataContext is MainWindowViewModel vm) vm.HideRefSelection();
            }
        }
    }

    private void OnFormulaBarGotFocus(object? sender, GotFocusEventArgs e) {
        if (sender is TextBox tb) {
            _activeFormulaEditor = tb;
            _isLastActionRefSelect = false; 
        }
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
                if (_activeFormulaEditor != null && _activeFormulaEditor.Text?.StartsWith("=") == true) {
                    string colName = MainWindowViewModel.GetColumnName(c);
                    string cellRef = $"{colName}{r + 1}";
                    string currentText = _activeFormulaEditor.Text ?? "";
                    
                    if (_isLastActionRefSelect && _lastInsertedRefIndex != -1) {
                         if (_lastInsertedRefIndex + _lastInsertedRefLength <= currentText.Length) {
                             currentText = currentText.Remove(_lastInsertedRefIndex, _lastInsertedRefLength);
                         } else {
                             _lastInsertedRefIndex = currentText.Length;
                         }
                         
                         currentText = currentText.Insert(_lastInsertedRefIndex, cellRef);
                         _activeFormulaEditor.Text = currentText;
                         _lastInsertedRefLength = cellRef.Length;
                         _activeFormulaEditor.CaretIndex = _lastInsertedRefIndex + _lastInsertedRefLength;

                    } else {
                        int caret = _activeFormulaEditor.CaretIndex;
                        _lastInsertedRefIndex = caret;
                        _activeFormulaEditor.Text = currentText.Insert(caret, cellRef);
                        _lastInsertedRefLength = cellRef.Length;
                        _activeFormulaEditor.CaretIndex = caret + cellRef.Length;
                        
                        _isLastActionRefSelect = true;
                    }
                    
                    vm.ShowRefSelection(r, c);
                    
                    _activeFormulaEditor.Opacity = 1;
                    _activeFormulaEditor.IsHitTestVisible = true;
                    
                    _activeFormulaEditor.Focus();
                    
                    e.Handled = true; 
                    return;
                }
                
                if (_activeFormulaEditor != null) {
                    _activeFormulaEditor.IsHitTestVisible = false;
                    _activeFormulaEditor.Opacity = 0;
                }

                _isLastActionRefSelect = false;
                _activeFormulaEditor = null;
                vm.HideRefSelection();

                panel.Focus();

                if (r == _anchorRowIndex && c == _anchorColIndex) {
                    EnterEditMode(panel, e.GetPosition(panel));
                    return;
                }

                _isSelecting = true;
                _anchorRowIndex = r;
                _anchorColIndex = c;
                
                vm.StartSelection(r, c);
                UpdateActiveCellVisual(vm);
                
                _capturedControl = panel;
                e.Pointer.Capture(panel);
            }
        }
    }
    
    private TextBox? EnterEditMode(Control parentControl, Point point) {
        var hit = parentControl.InputHitTest(point);
        
        if (hit is Visual visual) {
            var cellBorder = (visual as Border) ?? visual.FindAncestorOfType<Border>();
            if (cellBorder != null) {
                var textBox = cellBorder.FindDescendantOfType<TextBox>();
                if (textBox != null && !textBox.IsHitTestVisible) {
                    textBox.IsHitTestVisible = true;
                    textBox.Opacity = 1;
                    textBox.Focus();
                    _activeFormulaEditor = textBox; 
                    _isLastActionRefSelect = false;
                    return textBox;
                }
            }
        }
        return null;
    }

    private void OnCellEditorLostFocus(object? sender, RoutedEventArgs e) {
        if (sender is TextBox textBox) {
            textBox.IsHitTestVisible = false;
            textBox.Opacity = 0;
        }
    }

    private void OnPointerMoved(object? sender, PointerEventArgs e) {
        if (_isResizingColumn && _targetColumn != null) {
            var currentPosition = e.GetPosition(this);
            var delta = currentPosition.X - _lastMousePosition.X;
            _targetColumn.Width = Math.Max(30, _targetColumn.Width + delta);
            _lastMousePosition = currentPosition;
            
            if (DataContext is MainWindowViewModel vm) UpdateActiveCellVisual(vm);
            
        } else if (_isResizingRow && _targetRow != null) {
            var currentPosition = e.GetPosition(this);
            var delta = currentPosition.Y - _lastMousePosition.Y;
            _targetRow.Height = Math.Max(20, _targetRow.Height + delta);
            _lastMousePosition = currentPosition;

            if (DataContext is MainWindowViewModel vm) UpdateActiveCellVisual(vm);

        } else if (_isSelecting && DataContext is MainWindowViewModel vm) {
            var mousePoint = e.GetPosition(MainScrollViewer);
            double absX = mousePoint.X + MainScrollViewer.Offset.X;
            double absY = mousePoint.Y + MainScrollViewer.Offset.Y;

            var (r, c) = GetRowColumnAt(absX, absY, vm);
            if (r != -1 && c != -1) {
                vm.UpdateSelection(r, c);
            }
        }
    }

    private void OnPointerReleased(object? sender, PointerReleasedEventArgs e) {
        if (_isResizingColumn) { _isResizingColumn = false; _targetColumn = null; }
        if (_isResizingRow) { _isResizingRow = false; _targetRow = null; }
        if (_isSelecting) { _isSelecting = false; }
        if (_capturedControl != null) {
            e.Pointer.Capture(null);
            _capturedControl = null;
        }
    }
    
    private void UpdateActiveCellVisual(MainWindowViewModel vm) {
        if (_anchorRowIndex == -1 || _anchorColIndex == -1) {
            ActiveCellBorder.IsVisible = false;
            return;
        }
        var rect = GetCellRect(_anchorRowIndex, _anchorColIndex, vm);
        Canvas.SetLeft(ActiveCellBorder, rect.X);
        Canvas.SetTop(ActiveCellBorder, rect.Y);
        ActiveCellBorder.Width = rect.Width;
        ActiveCellBorder.Height = rect.Height;
        ActiveCellBorder.IsVisible = true;
    }

    private Rect GetCellRect(int rowIdx, int colIdx, MainWindowViewModel vm) {
        double x = 0;
        for (int i = 0; i < colIdx && i < vm.ColumnHeaders.Count; i++) x += vm.ColumnHeaders[i].Width;
        double y = 0;
        for (int i = 0; i < rowIdx && i < vm.Rows.Count; i++) y += vm.Rows[i].Height;
        
        double w = (colIdx >= 0 && colIdx < vm.ColumnHeaders.Count) ? vm.ColumnHeaders[colIdx].Width : 0;
        double h = (rowIdx >= 0 && rowIdx < vm.Rows.Count) ? vm.Rows[rowIdx].Height : 0;
        return new Rect(x, y, w, h);
    }

    private (int row, int col) GetRowColumnAt(double x, double y, MainWindowViewModel vm) {
        int colIndex = -1;
        double currentX = 0;
        for (int i = 0; i < vm.ColumnHeaders.Count; i++) {
            double w = vm.ColumnHeaders[i].Width;
            if (x >= currentX && x < currentX + w) { colIndex = i; break; }
            currentX += w;
        }
        int rowIndex = -1;
        double currentY = 0;
        for (int i = 0; i < vm.Rows.Count; i++) {
            double h = vm.Rows[i].Height;
            if (y >= currentY && y < currentY + h) { rowIndex = i; break; }
            currentY += h;
        }
        if (colIndex == -1 && x >= currentX && vm.ColumnHeaders.Count > 0) colIndex = vm.ColumnHeaders.Count - 1;
        if (rowIndex == -1 && y >= currentY && vm.Rows.Count > 0) rowIndex = vm.Rows.Count - 1;
        return (rowIndex, colIndex);
    }
    
    private async void OnOpenFileClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        var topLevel = Avalonia.Controls.TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = "Open Spreadsheet",
            AllowMultiple = false,
            FileTypeFilter = new[] {
                new FilePickerFileType("JSON Files") { Patterns = new[] { "*.json" } }
            }
        });

        if (files.Count >= 1 && DataContext is MainWindowViewModel vm) {
            vm.LoadData(files[0].Path.LocalPath);
        }
    }

    private async void OnSaveFileClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        var topLevel = Avalonia.Controls.TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = "Save Spreadsheet",
            DefaultExtension = "json",
            FileTypeChoices = new[] {
                new FilePickerFileType("JSON Files") { Patterns = new[] { "*.json" } }
            }
        });

        if (file != null && DataContext is MainWindowViewModel vm) {
            vm.SaveData(file.Path.LocalPath);
        }
    }
    
}