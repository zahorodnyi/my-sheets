using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.Models;

namespace MySheets.UI.ViewModels;

public partial class SheetViewModel : ObservableObject {
    public Worksheet Worksheet { get; }
    private int _anchorRow;
    private int _anchorCol;

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

        IsSelectionVisible = true;
        UpdateSelectionGeometry(row, col);
    }

    public void UpdateSelection(int row, int col) {
        UpdateSelectionGeometry(row, col);
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

    public void ShowRefSelection(int row, int col) {
        double x = 0;
        double y = 0;
        double w = ColumnHeaders[col].Width;
        double h = Rows[row].Height;

        for (int c = 0; c < col; c++) x += ColumnHeaders[c].Width;
        for (int r = 0; r < row; r++) y += Rows[r].Height;

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
}