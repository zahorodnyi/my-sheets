namespace MySheets.UI.ViewModels;

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using MySheets.Core.Models;

public class MainWindowViewModel : ObservableObject {
    private readonly Worksheet _sheet;

    public ObservableCollection<string> ColumnHeaders { get; }
    public ObservableCollection<RowViewModel> Rows { get; }

    public MainWindowViewModel() {
        _sheet = new Worksheet();
        ColumnHeaders = new ObservableCollection<string>();
        Rows = new ObservableCollection<RowViewModel>();

        const int RowCount = 50;
        const int ColCount = 25;

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