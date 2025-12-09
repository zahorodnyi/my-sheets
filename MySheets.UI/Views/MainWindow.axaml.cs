using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using MySheets.UI.ViewModels;
using Avalonia.VisualTree; 
using System.Linq;
using Avalonia.Input;
using System.Globalization;

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
                new FilePickerFileType("Supported Files") { Patterns = new[] { "*.json", "*.xlsx" } },
                new FilePickerFileType("JSON Files") { Patterns = new[] { "*.json" } },
                new FilePickerFileType("Excel Workbook") { Patterns = new[] { "*.xlsx" } }
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
            SuggestedFileName = "Sheet1",
            FileTypeChoices = new[] {
                new FilePickerFileType("JSON Files") { Patterns = new[] { "*.json" } },
                new FilePickerFileType("Excel Workbook") { Patterns = new[] { "*.xlsx" } }
            }
        });

        if (file != null && DataContext is MainWindowViewModel vm) {
            vm.SaveData(file.Path.LocalPath);
        }
    }

    private void OnInsertFunctionClick(object? sender, RoutedEventArgs e) {
        if (sender is Control control && control.Tag is string functionName) {
            if (DataContext is MainWindowViewModel vm && vm.ActiveSheet?.SelectedCell != null) {
                string formula = $"={functionName.ToUpper()}()";
                
                vm.ActiveSheet.SelectedCell.Expression = formula;

                var sheetView = this.GetVisualDescendants().OfType<SheetView>().FirstOrDefault();

                if (sheetView != null) {
                    sheetView.StartEditing(caretOffsetFromEnd: 1);
                }
            }
        }
    }

    private void OnFontSizeKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key == Key.Enter) {
            if (sender is TextBox tb && double.TryParse(tb.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double val)) {
                if (DataContext is MainWindowViewModel vm) {
                    vm.CurrentFontSize = val;
                    tb.Text = val.ToString(CultureInfo.InvariantCulture); 
                    e.Handled = true;
                    this.Focus(); 
                }
            }
        }
    }

    private void OnFontSizeLostFocus(object? sender, RoutedEventArgs e) {
        if (sender is TextBox tb && DataContext is MainWindowViewModel vm) {
            tb.Text = vm.CurrentFontSize.ToString(CultureInfo.InvariantCulture);
        }
    }
}