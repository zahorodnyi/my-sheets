namespace MySheets.UI.ViewModels;

using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using MySheets.Core.Models;

public class MainWindowViewModel : ObservableObject {
    private readonly Worksheet _sheet;

    public ObservableCollection<RowViewModel> Rows { get; }

    public MainWindowViewModel() {
        _sheet = new Worksheet();
        Rows = new ObservableCollection<RowViewModel>();

        const int RowCount = 10;
        const int ColCount = 5;

        for (var r = 0; r < RowCount; r++) {
            var rowCells = new List<CellViewModel>();
            for (var c = 0; c < ColCount; c++) {
                var cellModel = _sheet.GetCell(r, c);
                rowCells.Add(new CellViewModel(cellModel));
            }
            Rows.Add(new RowViewModel(rowCells));
        }
    }
}