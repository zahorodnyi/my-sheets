namespace MySheets.UI.ViewModels;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.Models;

public partial class MainWindowViewModel : ObservableObject {
    private readonly Worksheet _sheet;

    public ObservableCollection<string> ColumnHeaders { get; }
    public ObservableCollection<RowViewModel> Rows { get; }

    public MainWindowViewModel() {
        _sheet = new Worksheet();
        ColumnHeaders = new ObservableCollection<string>();
        Rows = new ObservableCollection<RowViewModel>();

        const int RowCount = 50; // todo move to utils
        const int ColCount = 25; // todo move to utils

        for (int c = 0; c < ColCount; c++) {
            ColumnHeaders.Add(GetColumnName(c));
        }

        for (var r = 0; r < RowCount; r++) {
            var rowCells = new List<CellViewModel>();
            for (var c = 0; c < ColCount; c++) {
                var cellModel = _sheet.GetCell(r, c);
                rowCells.Add(new CellViewModel(cellModel));
            }
            Rows.Add(new RowViewModel(rowCells, r + 1));
        }
    }

    [RelayCommand]
    private void AddRow() {
        int rowIndex = Rows.Count;
        var rowCells = new List<CellViewModel>();
        
        for (int c = 0; c < ColumnHeaders.Count; c++) {
            var cellModel = _sheet.GetCell(rowIndex, c);
            rowCells.Add(new CellViewModel(cellModel));
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
        ColumnHeaders.Add(GetColumnName(colIndex));

        for (int r = 0; r < Rows.Count; r++) {
            var cellModel = _sheet.GetCell(r, colIndex);
            Rows[r].Cells.Add(new CellViewModel(cellModel));
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