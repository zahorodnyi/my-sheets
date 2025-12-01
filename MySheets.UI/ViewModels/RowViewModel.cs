namespace MySheets.UI.ViewModels;

using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

public class RowViewModel : ObservableObject {
    public ObservableCollection<CellViewModel> Cells { get; }
    
    public string Header { get; }

    public RowViewModel(IEnumerable<CellViewModel> cells, int rowNumber) {
        Cells = new ObservableCollection<CellViewModel>(cells);
        Header = rowNumber.ToString();
    }
}