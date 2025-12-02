using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using MySheets.UI.ViewModels;

namespace MySheets.UI.Views;

public partial class MainWindow : Window {
    public static TextBox? GlobalFormulaBar { get; private set; }

    public MainWindow() {
        InitializeComponent();
        GlobalFormulaBar = this.FindControl<TextBox>("FormulaBar");
    }

    private void OnFormulaBarGotFocus(object? sender, global::Avalonia.Input.GotFocusEventArgs e) {
    }

    private async void OnOpenFileClick(object? sender, RoutedEventArgs e) {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var files = await topLevel.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = "Open Spreadsheet",
            AllowMultiple = false,
            FileTypeFilter = new[] {
                new FilePickerFileType("JSON Files") { Patterns = new[] { "*.json" } }
            }
        });

        if (files.Count >= 1 && DataContext is MainWindowViewModel vm) {
            vm.LoadData(files[0].Path.LocalPath);
        }
    }

    private async void OnSaveFileClick(object? sender, RoutedEventArgs e) {
        var topLevel = TopLevel.GetTopLevel(this);
        if (topLevel == null) return;

        var file = await topLevel.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = "Save Spreadsheet",
            DefaultExtension = "json",
            FileTypeChoices = new[] {
                new FilePickerFileType("JSON Files") { Patterns = new[] { "*.json" } }
            }
        });

        if (file != null && DataContext is MainWindowViewModel vm) {
            vm.SaveData(file.Path.LocalPath);
        }
    }
}