using CommunityToolkit.Mvvm.ComponentModel;

namespace MySheets.UI.ViewModels.SheetEditor;

public partial class ColumnViewModel : ObservableObject {
    [ObservableProperty]
    private string _header;

    [ObservableProperty]
    private double _width;
    
    [ObservableProperty] 
    private bool _isActive;

    public ColumnViewModel(string header, double width = 120) {
        _header = header;
        _width = width;
    }
}