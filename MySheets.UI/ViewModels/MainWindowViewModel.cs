using System;
using System.Collections.ObjectModel;
using System.Linq;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.Services;

namespace MySheets.UI.ViewModels;

public partial class MainWindowViewModel : ObservableObject {
    private readonly FileService _fileService;
    private int _untitledCount = 1;

    public ObservableCollection<SheetViewModel> Sheets { get; } = new();

    [ObservableProperty] 
    private SheetViewModel? _activeSheet;

    public MainWindowViewModel() {
        _fileService = new FileService();
        AddNewSheet();
    }

    [RelayCommand]
    public void AddNewSheet() {
        var newSheet = new SheetViewModel($"Sheet{_untitledCount++}");
        Sheets.Add(newSheet);
        ActiveSheet = newSheet;
    }

    [RelayCommand]
    public void CloseSheet(SheetViewModel sheet) {
        if (sheet == null) return;

        int index = Sheets.IndexOf(sheet);
        if (index == -1) return;

        bool wasActive = ActiveSheet == sheet;

        Sheets.Remove(sheet);

        if (wasActive && Sheets.Count > 0) {
            int newIndex = Math.Min(index, Sheets.Count - 1);
            ActiveSheet = Sheets[newIndex];
        }
    }

    [RelayCommand]
    private void AddRowToActive() {
        ActiveSheet?.AddRow();
    }

    [RelayCommand]
    private void RemoveRowFromActive() {
        ActiveSheet?.RemoveRow();
    }

    public void SaveData(string path) {
        if (ActiveSheet == null) return;
        _fileService.Save(path, ActiveSheet.Worksheet.Cells);
        ActiveSheet.Name = System.IO.Path.GetFileNameWithoutExtension(path);
    }

    public void LoadData(string path) {
        var data = _fileService.Load(path);
        
        var fileName = System.IO.Path.GetFileNameWithoutExtension(path);
        var newSheet = new SheetViewModel(fileName);
        
        foreach (var cellDto in data) {
            newSheet.Worksheet.SetCell(cellDto.Row, cellDto.Col, cellDto.Expression);
        }

        Sheets.Add(newSheet);
        ActiveSheet = newSheet;
    }

    public static string GetColumnName(int index) {
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