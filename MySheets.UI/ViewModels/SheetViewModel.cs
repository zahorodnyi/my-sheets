using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.Models;
using MySheets.Core.Utilities;

namespace MySheets.UI.ViewModels;

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

    private void UpdateAddressText() {
        string startCol = CellReferenceUtility.GetColumnName(_anchorCol);
        string startAddr = $"{startCol}{_anchorRow + 1}";

        if (_anchorRow == _currentRow && _anchorCol == _currentCol) {
            CurrentAddress = startAddr;
        } else {
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