using System;
using System.Collections.ObjectModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.IO;

namespace MySheets.UI.ViewModels;

public partial class MainWindowViewModel : ObservableObject {
    private readonly FileService _fileService;
    private int _untitledCount = 1;

    public ObservableCollection<SheetEditor.SheetViewModel> Sheets { get; } = new();

    [ObservableProperty] 
    private SheetEditor.SheetViewModel? _activeSheet;

    public MainWindowViewModel() {
        _fileService = new FileService();
        AddNewSheet();
    }

    [RelayCommand]
    public void AddNewSheet() {
        var newSheet = new SheetEditor.SheetViewModel($"Sheet{_untitledCount++}");
        Sheets.Add(newSheet);
        ActiveSheet = newSheet;
    }

    [RelayCommand]
    public void CloseSheet(SheetEditor.SheetViewModel sheet) {
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
    
    [RelayCommand]
    private void IncreaseFontSize() {
        ActiveSheet?.ApplyStyleToSelection(cell => {
             if (cell.FontSize < 72) cell.FontSize += 1;
        });
    }
    
    [RelayCommand]
    private void DecreaseFontSize() {
        ActiveSheet?.ApplyStyleToSelection(cell => {
             if (cell.FontSize > 6) cell.FontSize -= 1;
        });
    }
    
    [RelayCommand]
    private void ToggleBold() {
        ActiveSheet?.ApplyStyleToSelection(cell => {
            cell.IsBold = !cell.IsBold;
        });
    }

    [RelayCommand]
    private void ToggleItalic() {
        ActiveSheet?.ApplyStyleToSelection(cell => {
            cell.IsItalic = !cell.IsItalic;
        });
    }

    [RelayCommand]
    private void SetTextColor() {
        ActiveSheet?.ApplyStyleToSelection(cell => {
            string current = cell.Foreground?.ToString() ?? "";
            bool isRed = current.Contains("FF0000") || current.Contains("red", StringComparison.OrdinalIgnoreCase);
            cell.SetTextColor(isRed ? "#000000" : "#FF0000");
        });
    }

    [RelayCommand]
    private void SetFillColor() {
        ActiveSheet?.ApplyStyleToSelection(cell => {
            string current = cell.Background?.ToString() ?? "";
            bool isYellow = current.Contains("FFFF00") || current.Contains("yellow", StringComparison.OrdinalIgnoreCase);
            cell.SetBackgroundColor(isYellow ? "Transparent" : "#FFFF00");
        });
    }

    [RelayCommand]
    private void ToggleBorders() {
        ActiveSheet?.ApplyStyleToSelection(cell => {
            bool hasFullBorder = cell.BorderThickness.Left > 0;
            cell.SetBorder(hasFullBorder ? "0,0,1,1" : "1,1,1,1");
        });
    }

    public void SaveData(string path) {
        if (ActiveSheet == null) return;
        _fileService.Save(path, ActiveSheet.Worksheet.Cells);
        ActiveSheet.Name = System.IO.Path.GetFileNameWithoutExtension(path);
    }

    public void LoadData(string path) {
        var data = _fileService.Load(path);
        
        var fileName = System.IO.Path.GetFileNameWithoutExtension(path);
        var newSheet = new SheetEditor.SheetViewModel(fileName);
        
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