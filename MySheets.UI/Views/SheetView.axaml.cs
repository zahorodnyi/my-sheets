using System;
using System.Linq;
using System.Text.RegularExpressions;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.VisualTree;
using MySheets.UI.ViewModels;

namespace MySheets.UI.Views;

public partial class SheetView : UserControl {
    private bool _isResizingColumn;
    private bool _isResizingRow;
    private bool _isSelecting;
    
    private Point _lastMousePosition;
    private ColumnViewModel? _targetColumn;
    private RowViewModel? _targetRow;
    private Control? _capturedControl; 
    
    private int _anchorRowIndex = -1;
    private int _anchorColIndex = -1;

    private CellViewModel? _editingCellViewModel;
    private string _originalExpression = string.Empty;
    
    private bool _isLastActionRefSelect = false;
    private int _lastInsertedRefIndex = -1;
    private int _lastInsertedRefLength = 0;

    private readonly char[] _operators = new[] { '+', '-', '*', '/', '(', ')', '=', ',', '&' };

    public SheetView() {
        InitializeComponent();

        MainScrollViewer.ScrollChanged += (s, e) => {
            ColHeaderScrollViewer.Offset = new Vector(MainScrollViewer.Offset.X, 0);
            RowHeaderScrollViewer.Offset = new Vector(0, MainScrollViewer.Offset.Y);
        };

        CellPanel.AddHandler(PointerPressedEvent, OnCellPanelPointerPressed, RoutingStrategies.Tunnel);
        
        FloatingEditor.KeyDown += OnEditorKeyDown;
        FloatingEditor.LostFocus += OnEditorLostFocus;
    }

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e) {
        base.OnAttachedToVisualTree(e);
        if (MainWindow.GlobalFormulaBar != null) {
            MainWindow.GlobalFormulaBar.KeyDown += OnEditorKeyDown;
            MainWindow.GlobalFormulaBar.TextInput += OnEditorTextInput;
            MainWindow.GlobalFormulaBar.GotFocus += OnGlobalBarGotFocus;
        }
    }

    protected override void OnDetachedFromVisualTree(VisualTreeAttachmentEventArgs e) {
        base.OnDetachedFromVisualTree(e);
        if (MainWindow.GlobalFormulaBar != null) {
            MainWindow.GlobalFormulaBar.KeyDown -= OnEditorKeyDown;
            MainWindow.GlobalFormulaBar.TextInput -= OnEditorTextInput;
            MainWindow.GlobalFormulaBar.GotFocus -= OnGlobalBarGotFocus;
        }
    }

    private void OnGlobalBarGotFocus(object? sender, GotFocusEventArgs e) {
        if (ValidationPopup.IsVisible) return;

        if (DataContext is SheetViewModel vm && vm.SelectedCell != null) {
            if (_anchorRowIndex != -1 && _anchorColIndex != -1) {
                ActivateFloatingEditor(_anchorRowIndex, _anchorColIndex, vm, setFocus: false);
            }
        }
    }

    private void OnEditorTextInput(object? sender, TextInputEventArgs e) {
        if (ValidationPopup.IsVisible) return;

        _isLastActionRefSelect = false;
        if (DataContext is SheetViewModel vm && vm.IsRefSelectionVisible) {
             vm.HideRefSelection();
        }
    }

    private void OnEditorKeyDown(object? sender, KeyEventArgs e) {
        if (ValidationPopup.IsVisible) return;

        if (e.Key == Key.Enter) {
             if (TryCommitEditor()) { 
                 this.Focus(); 
                 e.Handled = true;
                 
                 if (DataContext is SheetViewModel vm && _anchorRowIndex != -1) {
                     int nextRow = Math.Min(_anchorRowIndex + 1, vm.Rows.Count - 1);
                     vm.StartSelection(nextRow, _anchorColIndex);
                     _anchorRowIndex = nextRow;
                     UpdateActiveCellVisual(vm);
                 }
             } else {
                 e.Handled = true; 
             }
             return;
        }
        
        if (e.Key == Key.Escape) {
            CancelEditor();
            e.Handled = true;
            return;
        }
    }

    private TextBox? GetActiveEditor() {
        if (FloatingEditor.IsVisible && FloatingEditor.IsFocused) return FloatingEditor;
        
        var globalBar = MainWindow.GlobalFormulaBar;
        if (globalBar != null) {
            if (globalBar.IsFocused) return globalBar;
            if (FloatingEditor.IsVisible) return FloatingEditor;
        }
        return null;
    }

    protected override void OnTextInput(TextInputEventArgs e) {
        base.OnTextInput(e);
        if (ValidationPopup.IsVisible) return;
        
        if (DataContext is SheetViewModel vm && vm.IsRefSelectionVisible) {
             _isLastActionRefSelect = false;
             vm.HideRefSelection();
        }

        var activeEditor = GetActiveEditor();
        if (activeEditor != null && activeEditor.IsFocused) return; 

        if (!string.IsNullOrEmpty(e.Text) && 
            _anchorRowIndex != -1 && 
            _anchorColIndex != -1 && 
            DataContext is SheetViewModel viewModel) {
            
            ActivateFloatingEditor(_anchorRowIndex, _anchorColIndex, viewModel, setFocus: true);
            
            if (FloatingEditor.IsVisible) {
                FloatingEditor.Text = e.Text;
                FloatingEditor.CaretIndex = e.Text.Length;
            }
        }
    }
    
    protected override void OnKeyDown(KeyEventArgs e) {
        base.OnKeyDown(e);
        if (ValidationPopup.IsVisible) return;
        
        bool isEditing = (FloatingEditor.IsVisible && FloatingEditor.IsFocused) || 
                         (MainWindow.GlobalFormulaBar != null && MainWindow.GlobalFormulaBar.IsFocused);

        if (!isEditing && !FloatingEditor.IsVisible && DataContext is SheetViewModel vm) {
             int r = _anchorRowIndex;
             int c = _anchorColIndex;
             bool moved = false;
             
             if (e.Key == Key.Up) { r = Math.Max(0, r - 1); moved = true; }
             else if (e.Key == Key.Down) { r = Math.Min(vm.Rows.Count - 1, r + 1); moved = true; }
             else if (e.Key == Key.Left) { c = Math.Max(0, c - 1); moved = true; }
             else if (e.Key == Key.Right) { c = Math.Min(vm.ColumnHeaders.Count - 1, c + 1); moved = true; }
             
             if (moved) {
                 _anchorRowIndex = r;
                 _anchorColIndex = c;
                 vm.StartSelection(r, c);
                 UpdateActiveCellVisual(vm);
                 e.Handled = true;
             }
        }
    }

    private void OnCellPanelPointerPressed(object? sender, PointerPressedEventArgs e) {
        if (ValidationPopup.IsVisible) { e.Handled = true; return; }

        var properties = e.GetCurrentPoint(this).Properties;
        if (!properties.IsLeftButtonPressed) return;
        if (_isResizingColumn || _isResizingRow) return;

        var activeEditor = GetActiveEditor();

        if (DataContext is SheetViewModel vm && sender is Control panel) {
            var mousePoint = e.GetPosition(MainScrollViewer);
            double absX = mousePoint.X + MainScrollViewer.Offset.X;
            double absY = mousePoint.Y + MainScrollViewer.Offset.Y;
            var (r, c) = GetRowColumnAt(absX, absY, vm);
            
            if (r != -1 && c != -1) {
                bool isSelfClick = (r == _anchorRowIndex && c == _anchorColIndex);

                if (!isSelfClick && activeEditor != null && activeEditor.Text?.StartsWith("=") == true) {
                    if (ShouldPickReference(activeEditor)) {
                        string colName = MainWindowViewModel.GetColumnName(c);
                        string cellRef = $"{colName}{r + 1}";
                        InsertReferenceIntoEditor(activeEditor, cellRef);
                        
                        vm.ShowRefSelection(r, c);
                        
                        if (activeEditor == FloatingEditor) {
                            activeEditor.Focus();
                        } else if (MainWindow.GlobalFormulaBar != null && MainWindow.GlobalFormulaBar.IsFocused) {
                             MainWindow.GlobalFormulaBar.Focus();
                        }
                        
                        e.Handled = true; 
                        return;
                    }
                }
                
                if (!isSelfClick) {
                    if (!TryCommitEditor()) {
                        e.Handled = true; 
                        return; 
                    }
                }

                _isLastActionRefSelect = false;
                vm.HideRefSelection();

                this.Focus();

                if (isSelfClick) {
                    ActivateFloatingEditor(r, c, vm, setFocus: true);
                    e.Handled = true;
                    return;
                }

                _isSelecting = true;
                _anchorRowIndex = r;
                _anchorColIndex = c;
                
                vm.StartSelection(r, c);
                UpdateActiveCellVisual(vm);
                _capturedControl = panel;
                e.Pointer.Capture(panel);

                if (e.ClickCount == 2) {
                    ActivateFloatingEditor(r, c, vm, setFocus: true);
                }
            }
        }
    }

    private bool ShouldPickReference(TextBox editor) {
        string text = editor.Text ?? "";
        int caret = editor.CaretIndex;
        
        var existingRef = GetReferenceAtCaret(editor);
        if (existingRef.length > 0) return true;

        if (caret >= text.Length) {
             if (text.Length == 0) return false;
             char lastChar = text[text.Length - 1];
             if (_operators.Contains(lastChar)) return true;
             if (lastChar == '=') return true;
             if (char.IsDigit(lastChar)) return false;
             if (char.IsWhiteSpace(lastChar)) return false;
             return false;
        }

        if (caret > 0) {
            char charBefore = text[caret - 1];
            if (_operators.Contains(charBefore) || charBefore == '=') return true;
        }

        return false;
    }

    private void InsertReferenceIntoEditor(TextBox activeEditor, string cellRef) {
        string currentText = activeEditor.Text ?? "";
        
        if (_isLastActionRefSelect && _lastInsertedRefIndex != -1) {
             if (_lastInsertedRefIndex + _lastInsertedRefLength <= currentText.Length) {
                 currentText = currentText.Remove(_lastInsertedRefIndex, _lastInsertedRefLength);
             } 
             else {
                 _lastInsertedRefIndex = currentText.Length;
             }
             currentText = currentText.Insert(_lastInsertedRefIndex, cellRef);
             activeEditor.Text = currentText;
             _lastInsertedRefLength = cellRef.Length;
             activeEditor.CaretIndex = _lastInsertedRefIndex + _lastInsertedRefLength;
        } 
        else {
            var existingRef = GetReferenceAtCaret(activeEditor);
            if (existingRef.length > 0) {
                currentText = currentText.Remove(existingRef.start, existingRef.length);
                currentText = currentText.Insert(existingRef.start, cellRef);
                activeEditor.Text = currentText;
                
                _lastInsertedRefIndex = existingRef.start;
                _lastInsertedRefLength = cellRef.Length;
                activeEditor.CaretIndex = existingRef.start + cellRef.Length;
                _isLastActionRefSelect = true;
            } else {
                int caret = activeEditor.CaretIndex; 
                _lastInsertedRefIndex = caret;
                activeEditor.Text = currentText.Insert(caret, cellRef);
                _lastInsertedRefLength = cellRef.Length;
                activeEditor.CaretIndex = caret + cellRef.Length;
                _isLastActionRefSelect = true;
            }
        }
    }

    private (int start, int length, string text) GetReferenceAtCaret(TextBox editor) {
        string text = editor.Text ?? "";
        int caret = editor.CaretIndex;
        if (caret > text.Length) caret = text.Length;

        int start = caret;
        while (start > 0) {
            char c = text[start - 1];
            if (_operators.Contains(c) || char.IsWhiteSpace(c) || c == '=') break;
            start--;
        }
        
        int end = caret;
        while (end < text.Length) {
            char c = text[end];
            if (_operators.Contains(c) || char.IsWhiteSpace(c) || c == '=') break;
            end++;
        }
        
        if (end <= start) return (-1, 0, "");
        
        string potentialRef = text.Substring(start, end - start);
        if (Regex.IsMatch(potentialRef, @"^[A-Za-z]+[0-9]+$")) {
            return (start, end - start, potentialRef);
        }
        return (-1, 0, "");
    }
    
    private void ActivateFloatingEditor(int r, int c, SheetViewModel vm, bool setFocus) {
        if (r < 0 || r >= vm.Rows.Count) return;
        var row = vm.Rows[r];
        if (c < 0 || c >= row.Cells.Count) return;
        var cell = row.Cells[c];

        _editingCellViewModel = cell;
        _originalExpression = cell.Expression;

        var rect = GetCellRect(r, c, vm);

        Canvas.SetLeft(FloatingEditor, rect.X);
        Canvas.SetTop(FloatingEditor, rect.Y);
        
        FloatingEditor.Width = rect.Width;
        FloatingEditor.Height = rect.Height; 

        FloatingEditor.Text = cell.Expression;
        FloatingEditor.IsVisible = true;
        
        if (MainWindow.GlobalFormulaBar != null) {
            MainWindow.GlobalFormulaBar.Text = cell.Expression;
        }
        
        if (setFocus) {
            FloatingEditor.Focus();
            FloatingEditor.CaretIndex = FloatingEditor.Text?.Length ?? 0;
        }
        
        _isLastActionRefSelect = false;
    }

    private bool TryCommitEditor() {
        if (!FloatingEditor.IsVisible && (MainWindow.GlobalFormulaBar == null || !MainWindow.GlobalFormulaBar.IsFocused)) {
            return true;
        }

        string newExpression = FloatingEditor.IsVisible ? FloatingEditor.Text ?? "" : MainWindow.GlobalFormulaBar?.Text ?? "";

        if (newExpression.StartsWith("=")) {
            if (DataContext is SheetViewModel vm) {
                if (!vm.ValidateFormula(newExpression)) {
                    ShowValidationError();
                    return false;
                }
            }
        }

        if (_editingCellViewModel != null) {
            _editingCellViewModel.Expression = newExpression;
            _editingCellViewModel = null;
        }

        FloatingEditor.IsVisible = false;
        return true;
    }

    private void CancelEditor() {
        if (ValidationPopup.IsVisible) {
            HideValidationError();
            return;
        }

        if (_editingCellViewModel != null) {
            if (MainWindow.GlobalFormulaBar != null) {
                MainWindow.GlobalFormulaBar.Text = _originalExpression;
            }
            
            FloatingEditor.IsVisible = false;
            _editingCellViewModel = null;
            this.Focus();
        }
        else if (MainWindow.GlobalFormulaBar != null && MainWindow.GlobalFormulaBar.IsFocused) {
            MainWindow.GlobalFormulaBar.Text = _originalExpression;
            this.Focus();
        }
    }

    private void OnEditorLostFocus(object? sender, RoutedEventArgs e) {
        if (MainWindow.GlobalFormulaBar != null && MainWindow.GlobalFormulaBar.IsFocused) return;
        
        if (ValidationPopup.IsVisible) return;

        if (FloatingEditor.IsVisible) {
             TryCommitEditor();
        }
    }

    private void ShowValidationError() {
        ValidationPopup.IsVisible = true;
    }

    private void HideValidationError() {
        ValidationPopup.IsVisible = false;
    }

    private void OnValidationOkClick(object? sender, RoutedEventArgs e) {
        HideValidationError();
        
        if (FloatingEditor.IsVisible) {
            FloatingEditor.Focus();
        } else if (MainWindow.GlobalFormulaBar != null && MainWindow.GlobalFormulaBar.IsFocused) {
            MainWindow.GlobalFormulaBar.Focus();
        }
    }

    private void OnPointerMoved(object? sender, PointerEventArgs e) {
        if (DataContext is not SheetViewModel vm) return;
        if (ValidationPopup.IsVisible) return;

        if (_isResizingColumn && _targetColumn != null) {
            var currentPosition = e.GetPosition(this);
            var delta = currentPosition.X - _lastMousePosition.X;
            _targetColumn.Width = Math.Max(30, _targetColumn.Width + delta);
            _lastMousePosition = currentPosition;
            UpdateActiveCellVisual(vm);
            
        } 
        else if (_isResizingRow && _targetRow != null) {
            var currentPosition = e.GetPosition(this);
            var delta = currentPosition.Y - _lastMousePosition.Y;
            _targetRow.Height = Math.Max(20, _targetRow.Height + delta);
            _lastMousePosition = currentPosition;
            UpdateActiveCellVisual(vm);

        } 
        else if (_isSelecting) {
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
    
    private void OnColumnHeaderPressed(object? sender, PointerPressedEventArgs e) {
        if (ValidationPopup.IsVisible) return;
        TryCommitEditor();
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
        if (ValidationPopup.IsVisible) return;
        TryCommitEditor();
        if (sender is Border border && border.Tag is RowViewModel row) {
            _isResizingRow = true;
            _targetRow = row;
            _lastMousePosition = e.GetPosition(this);
            _capturedControl = border;
            e.Pointer.Capture(border);
            e.Handled = true;
        }
    }

    private void UpdateActiveCellVisual(SheetViewModel vm) {
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

    private Rect GetCellRect(int rowIdx, int colIdx, SheetViewModel vm) {
        double x = 0;
        for (int i = 0; i < colIdx && i < vm.ColumnHeaders.Count; i++) x += vm.ColumnHeaders[i].Width;
        double y = 0;
        for (int i = 0; i < rowIdx && i < vm.Rows.Count; i++) y += vm.Rows[i].Height;
        
        double w = (colIdx >= 0 && colIdx < vm.ColumnHeaders.Count) ? vm.ColumnHeaders[colIdx].Width : 0;
        double h = (rowIdx >= 0 && rowIdx < vm.Rows.Count) ? vm.Rows[rowIdx].Height : 0;
        return new Rect(x, y, w, h);
    }

    private (int row, int col) GetRowColumnAt(double x, double y, SheetViewModel vm) {
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
}