using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.Common;
using MySheets.Core.Domain;

namespace MySheets.UI.ViewModels.SheetEditor;

public partial class SheetViewModel : ObservableObject {
    public Worksheet Worksheet { get; }
    private int _anchorRow;
    private int _anchorCol;
    private int _currentRow;
    private int _currentCol;

    [ObservableProperty] private string _name;
    
    [ObservableProperty] private double _selectionX;
    [ObservableProperty] private double _selectionY;
    [ObservableProperty] private double _selectionWidth;
    [ObservableProperty] private double _selectionHeight;
    [ObservableProperty] private bool _isSelectionVisible;

    [ObservableProperty] private double _refSelectionX;
    [ObservableProperty] private double _refSelectionY;
    [ObservableProperty] private double _refSelectionWidth;
    [ObservableProperty] private double _refSelectionHeight;
    [ObservableProperty] private bool _isRefSelectionVisible;

    [ObservableProperty] private CellViewModel? _selectedCell;
    [ObservableProperty] private string _currentAddress = "A1";

    private RowViewModel? _activeRowHeader;
    private ColumnViewModel? _activeColHeader;

    public ObservableCollection<ColumnViewModel> ColumnHeaders { get; }
    public ObservableCollection<RowViewModel> Rows { get; }

    public SheetViewModel(string name) {
        Name = name;
        Worksheet = new Worksheet();
        ColumnHeaders = new ObservableCollection<ColumnViewModel>();
        Rows = new ObservableCollection<RowViewModel>();

        const int RowCount = 50;
        const int ColCount = 25;

        for (int c = 0; c < ColCount; c++) {
            ColumnHeaders.Add(new ColumnViewModel(MainWindowViewModel.GetColumnName(c)));
        }

        for (var r = 0; r < RowCount; r++) {
            var rowCells = new List<CellViewModel>();
            for (var c = 0; c < ColCount; c++) {
                var cellModel = Worksheet.GetCell(r, c);
                rowCells.Add(new CellViewModel(cellModel, Worksheet, ColumnHeaders[c]));
            }
            Rows.Add(new RowViewModel(rowCells, r + 1));
        }

        Worksheet.CellStateChanged += OnCellStateChanged;
    }

    private void OnCellStateChanged(int row, int col) {
        if (row < Rows.Count && col < Rows[row].Cells.Count) {
             var vm = Rows[row].Cells[col];
             vm.Refresh();
        }
    }

    public void StartSelection(int row, int col) {
        _anchorRow = row;
        _anchorCol = col;
        
        if (row >= 0 && row < Rows.Count && col >= 0 && col < Rows[row].Cells.Count) {
            SelectedCell = Rows[row].Cells[col];
        }
        
        if (_activeRowHeader != null) _activeRowHeader.IsActive = false;
        if (_activeColHeader != null) _activeColHeader.IsActive = false;

        if (row >= 0 && row < Rows.Count) {
            _activeRowHeader = Rows[row];
            _activeRowHeader.IsActive = true;
        }
        if (col >= 0 && col < ColumnHeaders.Count) {
            _activeColHeader = ColumnHeaders[col];
            _activeColHeader.IsActive = true;
        }

        IsSelectionVisible = true;
        UpdateSelection(row, col);
    }

    public void UpdateSelection(int row, int col) {
        if (row < 0 || col < 0) return;
        _currentRow = row;
        _currentCol = col;
        UpdateSelectionGeometry(row, col);
        UpdateAddressText();
    }

    public void ApplyStyleToSelection(Action<CellViewModel> applyAction) {
        int r1 = Math.Min(_anchorRow, _currentRow);
        int r2 = Math.Max(_anchorRow, _currentRow);
        int c1 = Math.Min(_anchorCol, _currentCol);
        int c2 = Math.Max(_anchorCol, _currentCol);

        for (int r = r1; r <= r2; r++) {
            if (r >= Rows.Count) continue;
            var row = Rows[r];
            for (int c = c1; c <= c2; c++) {
                if (c >= row.Cells.Count) continue;
                applyAction(row.Cells[c]);
            }
        }
    }

    public void SetFontSizeForSelection(double newSize) {
        int r1 = Math.Min(_anchorRow, _currentRow);
        int r2 = Math.Max(_anchorRow, _currentRow);
        int c1 = Math.Min(_anchorCol, _currentCol);
        int c2 = Math.Max(_anchorCol, _currentCol);

        double requiredHeight = newSize * 1.4;

        for (int r = r1; r <= r2; r++) {
            if (r >= Rows.Count) continue;
            
            var row = Rows[r];

            if (row.IsHeightManual) {
                if (row.Height < requiredHeight) {
                    row.Height = requiredHeight;
                }
            } 
            else {
                row.Height = Math.Max(25, requiredHeight);
            }

            for (int c = c1; c <= c2; c++) {
                if (c >= row.Cells.Count) continue;
                row.Cells[c].FontSize = newSize;
            }
        }
    }
    
    public void ApplyBorderSelection(string mode) {
        int r1 = Math.Min(_anchorRow, _currentRow);
        int r2 = Math.Max(_anchorRow, _currentRow);
        int c1 = Math.Min(_anchorCol, _currentCol);
        int c2 = Math.Max(_anchorCol, _currentCol);

        for (int r = r1; r <= r2; r++) {
            if (r >= Rows.Count) continue;
            var row = Rows[r];
            
            for (int c = c1; c <= c2; c++) {
                if (c >= row.Cells.Count) continue;
                var cell = row.Cells[c];

                if (mode == "None") {
                    cell.SetBorder("0,0,1,1");
                }
                else if (mode == "All") {
                    int left = (c == c1) ? 1 : 0;
                    int top = (r == r1) ? 1 : 0;
                    
                    cell.SetBorder($"{left},{top},1.00,1.00");
                }
                else if (mode == "Outside") {
                    bool isLeftEdge = (c == c1);
                    bool isTopEdge = (r == r1);
                    bool isRightEdge = (c == c2);
                    bool isBottomEdge = (r == r2);

                    if (!isLeftEdge && !isTopEdge && !isRightEdge && !isBottomEdge) continue;

                    var (curL, curT, curR, curB) = GetCurrentBorderValues(cell);

                    string newL = isLeftEdge ? "1" : (curL > 0 ? "1.00" : "0");
                    string newT = isTopEdge ? "1" : (curT > 0 ? "1.00" : "0");
                    string newR = isRightEdge ? "1.00" : (curR > 0 ? "1.00" : "0"); 
                    string newB = isBottomEdge ? "1.00" : (curB > 0 ? "1.00" : "0");
                    
                    cell.SetBorder($"{newL},{newT},{newR},{newB}");
                }
            }
        }
    }

    private (int, int, int, int) GetCurrentBorderValues(CellViewModel vm) {
        var parts = vm.BorderThickness.ToString().Split(',');
        if (parts.Length == 4) {
            return ((int)vm.BorderThickness.Left, (int)vm.BorderThickness.Top, (int)vm.BorderThickness.Right, (int)vm.BorderThickness.Bottom);
        }
        return (0, 0, 1, 1);
    }

    private void UpdateAddressText() {
        string startCol = CellReferenceUtility.GetColumnName(_anchorCol);
        string startAddr = $"{startCol}{_anchorRow + 1}";

        if (_anchorRow == _currentRow && _anchorCol == _currentCol) {
            CurrentAddress = startAddr;
        } 
        else {
            int r1 = Math.Min(_anchorRow, _currentRow);
            int r2 = Math.Max(_anchorRow, _currentRow);
            int c1 = Math.Min(_anchorCol, _currentCol);
            int c2 = Math.Max(_anchorCol, _currentCol);

            string tlCol = CellReferenceUtility.GetColumnName(c1);
            string brCol = CellReferenceUtility.GetColumnName(c2);
            CurrentAddress = $"{tlCol}{r1 + 1}:{brCol}{r2 + 1}";
        }
    }

    private void UpdateSelectionGeometry(int currentRow, int currentCol) {
        int startR = Math.Min(_anchorRow, currentRow);
        int endR = Math.Max(_anchorRow, currentRow);
        int startC = Math.Min(_anchorCol, currentCol);
        int endC = Math.Max(_anchorCol, currentCol);

        double x = 0;
        double y = 0;
        double w = 0;
        double h = 0;

        for (int c = 0; c < startC; c++) x += ColumnHeaders[c].Width;
        for (int r = 0; r < startR; r++) y += Rows[r].Height;
        for (int c = startC; c <= endC; c++) w += ColumnHeaders[c].Width;
        for (int r = startR; r <= endR; r++) h += Rows[r].Height;

        SelectionX = x;
        SelectionY = y;
        SelectionWidth = w;
        SelectionHeight = h;
    }

    public void ShowRefSelection(int row, int col, int endRow = -1, int endCol = -1) {
         if (endRow == -1) endRow = row;
         if (endCol == -1) endCol = col;
         
         int r1 = Math.Min(row, endRow);
         int r2 = Math.Max(row, endRow);
         int c1 = Math.Min(col, endCol);
         int c2 = Math.Max(col, endCol);

         double x = 0, y = 0, w = 0, h = 0;
         
         for (int c = 0; c < c1; c++) x += ColumnHeaders[c].Width;
         for (int r = 0; r < r1; r++) y += Rows[r].Height;
         for (int c = c1; c <= c2; c++) w += ColumnHeaders[c].Width;
         for (int r = r1; r <= r2; r++) h += Rows[r].Height;

         RefSelectionX = x;
         RefSelectionY = y;
         RefSelectionWidth = w;
         RefSelectionHeight = h;
         IsRefSelectionVisible = true;
    }

    public void HideRefSelection() {
        IsRefSelectionVisible = false;
    }

    [RelayCommand]
    public void AddRow() {
        int rowIndex = Rows.Count;
        var rowCells = new List<CellViewModel>();
        for (int c = 0; c < ColumnHeaders.Count; c++) {
            var cellModel = Worksheet.GetCell(rowIndex, c);
            rowCells.Add(new CellViewModel(cellModel, Worksheet, ColumnHeaders[c]));
        }
        Rows.Add(new RowViewModel(rowCells, rowIndex + 1));
    }

    [RelayCommand]
    public void RemoveRow() {
        if (Rows.Count > 0) Rows.RemoveAt(Rows.Count - 1);
    }

    [RelayCommand]
    public void AddColumn() {
        int colIndex = ColumnHeaders.Count;
        var newColumn = new ColumnViewModel(MainWindowViewModel.GetColumnName(colIndex));
        ColumnHeaders.Add(newColumn);
        for (int r = 0; r < Rows.Count; r++) {
            var cellModel = Worksheet.GetCell(r, colIndex);
            Rows[r].Cells.Add(new CellViewModel(cellModel, Worksheet, newColumn));
        }
    }

    [RelayCommand]
    public void RemoveColumn() {
        if (ColumnHeaders.Count > 0) {
            int lastColIndex = ColumnHeaders.Count - 1;
            ColumnHeaders.RemoveAt(lastColIndex);
            foreach (var row in Rows) {
                if (row.Cells.Count > 0) row.Cells.RemoveAt(row.Cells.Count - 1);
            }
        }
    }
    
    public bool ValidateFormula(string formula) {
        return Worksheet.IsFormulaValid(formula);
    }
}