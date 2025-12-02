namespace MySheets.UI.ViewModels;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.Models;

public partial class MainWindowViewModel : ObservableObject {
    private readonly Worksheet _sheet;
    private int _anchorRow;
    private int _anchorCol;

    [ObservableProperty] private double _selectionX;
    [ObservableProperty] private double _selectionY;
    [ObservableProperty] private double _selectionWidth;
    [ObservableProperty] private double _selectionHeight;
    [ObservableProperty] private bool _isSelectionVisible;

    public ObservableCollection<ColumnViewModel> ColumnHeaders { get; }
    public ObservableCollection<RowViewModel> Rows { get; }

    public MainWindowViewModel() {
        _sheet = new Worksheet();
        ColumnHeaders = new ObservableCollection<ColumnViewModel>();
        Rows = new ObservableCollection<RowViewModel>();

        const int RowCount = 50;
        const int ColCount = 25;

        for (int c = 0; c < ColCount; c++) {
            ColumnHeaders.Add(new ColumnViewModel(GetColumnName(c)));
        }

        for (var r = 0; r < RowCount; r++) {
            var rowCells = new List<CellViewModel>();
            for (var c = 0; c < ColCount; c++) {
                var cellModel = _sheet.GetCell(r, c);
                rowCells.Add(new CellViewModel(cellModel, ColumnHeaders[c]));
            }
            Rows.Add(new RowViewModel(rowCells, r + 1));
        }
    }

    public void StartSelection(int row, int col) {
        _anchorRow = row;
        _anchorCol = col;
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

        for (int c = 0; c < startC; c++) {
            x += ColumnHeaders[c].Width;
        }

        for (int r = 0; r < startR; r++) {
            y += Rows[r].Height;
        }

        for (int c = startC; c <= endC; c++) {
            w += ColumnHeaders[c].Width;
        }

        for (int r = startR; r <= endR; r++) {
            h += Rows[r].Height;
        }

        SelectionX = x;
        SelectionY = y;
        SelectionWidth = w;
        SelectionHeight = h;
    }

    [RelayCommand]
    private void AddRow() {
        int rowIndex = Rows.Count;
        var rowCells = new List<CellViewModel>();
        
        for (int c = 0; c < ColumnHeaders.Count; c++) {
            var cellModel = _sheet.GetCell(rowIndex, c);
            rowCells.Add(new CellViewModel(cellModel, ColumnHeaders[c]));
        }
        
        Rows.Add(new RowViewModel(rowCells, rowIndex + 1));
    }

    [RelayCommand]
    private void RemoveRow() {
        if (Rows.Count > 0) {
            Rows.RemoveAt(Rows.Count - 1);
        }
    }

    [RelayCommand]
    private void AddColumn() {
        int colIndex = ColumnHeaders.Count;
        var newColumn = new ColumnViewModel(GetColumnName(colIndex));
        ColumnHeaders.Add(newColumn);

        for (int r = 0; r < Rows.Count; r++) {
            var cellModel = _sheet.GetCell(r, colIndex);
            Rows[r].Cells.Add(new CellViewModel(cellModel, newColumn));
        }
    }

    [RelayCommand]
    private void RemoveColumn() {
        if (ColumnHeaders.Count > 0) {
            int lastColIndex = ColumnHeaders.Count - 1;
            ColumnHeaders.RemoveAt(lastColIndex);

            foreach (var row in Rows) {
                if (row.Cells.Count > 0) {
                    row.Cells.RemoveAt(row.Cells.Count - 1);
                }
            }
        }
    }

    private static string GetColumnName(int index) {
        int dividend = index + 1;
        string columnName = string.Empty;
        while (dividend > 0) {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
            dividend = (int)((dividend - modulo) / 26);
        }
        return columnName;
    }
}