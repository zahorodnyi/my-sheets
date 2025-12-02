namespace MySheets.UI.ViewModels;

using CommunityToolkit.Mvvm.ComponentModel;

public partial class ColumnViewModel : ObservableObject {
    [ObservableProperty]
    private string _header;

    [ObservableProperty]
    private double _width;

    public ColumnViewModel(string header, double width = 120) {
        _header = header;
        _width = width;
    }
}