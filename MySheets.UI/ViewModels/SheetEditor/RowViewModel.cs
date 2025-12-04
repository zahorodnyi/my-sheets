using System.Collections.Generic;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;

namespace MySheets.UI.ViewModels.SheetEditor;

public partial class RowViewModel : ObservableObject {
    public ObservableCollection<CellViewModel> Cells { get; }
    
    public string Header { get; }

    [ObservableProperty]
    private double _height;
    
    [ObservableProperty] 
    private bool _isActive;

    public RowViewModel(IEnumerable<CellViewModel> cells, int rowNumber, double height = 25) {
        Cells = new ObservableCollection<CellViewModel>(cells);
        Header = rowNumber.ToString();
        _height = height;
    }
}