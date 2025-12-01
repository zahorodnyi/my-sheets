namespace MySheets.UI.ViewModels;

using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

public class RowViewModel : ObservableObject {
    public ObservableCollection<CellViewModel> Cells { get; }

    public RowViewModel(IEnumerable<CellViewModel> cells) {
        Cells = new ObservableCollection<CellViewModel>(cells);
    }
}