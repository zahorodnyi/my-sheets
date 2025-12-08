using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Avalonia.Layout; 
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


    public List<string> TextColors { get; } = new() {
        "#000000", // Black
        "#434343", // Dark Grey
        "#666666", // Grey
        "#999999", // Light Grey
        "#FFFFFF", // White
        "#CC0000", // Dark Red
        "#E69138", // Dark Orange
        "#F1C232", // Dark Yellow
        "#6AA84F", // Dark Green
        "#3D85C6", // Dark Cyan/Blue
        "#073763", // Navy
        "#674EA7"  // Purple
    };

    public List<string> FillColors { get; } = new() {
        "#FFFFFF", // White
        "#F3F3F3", // Light Grey
        "#B7B7B7", // Grey
        "#F4CCCC", // Pastel Red
        "#FCE5CD", // Pastel Orange
        "#FFF2CC", // Pastel Yellow
        "#D9EAD3", // Pastel Green
        "#D0E0E3", // Pastel Cyan
        "#CFE2F3", // Pastel Blue
        "#D9D2E9", // Pastel Purple
        "#EAD1DC"  // Pastel Pink
    };

    [ObservableProperty]
    private string _currentTextColor = "#000000";

    [ObservableProperty]
    private string _currentFillColor = "Transparent";


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
    private void ApplyTextColor(string? colorHex) {
        if (colorHex != null) {
            CurrentTextColor = colorHex;
        }
        
        ActiveSheet?.ApplyStyleToSelection(cell => {
            cell.SetTextColor(CurrentTextColor);
        });
    }

    [RelayCommand]
    private void ResetTextColor() {
        CurrentTextColor = "#000000";
        ActiveSheet?.ApplyStyleToSelection(cell => {
            cell.SetTextColor("#000000");
        });
    }

    [RelayCommand]
    private void ApplyFillColor(string? colorHex) {
        if (colorHex != null) {
            CurrentFillColor = colorHex;
        }

        ActiveSheet?.ApplyStyleToSelection(cell => {
            cell.SetBackgroundColor(CurrentFillColor);
        });
    }

    [RelayCommand]
    private void ResetFillColor() {
        CurrentFillColor = "Transparent";
        ActiveSheet?.ApplyStyleToSelection(cell => {
            cell.SetBackgroundColor("Transparent");
        });
    }

    [RelayCommand]
    private void ToggleBorders() {
        ActiveSheet?.ApplyStyleToSelection(cell => {
            bool hasFullBorder = cell.BorderThickness.Left > 0;
            cell.SetBorder(hasFullBorder ? "0,0,1,1" : "1,1,1,1");
        });
    }

    [RelayCommand]
    private void SetAlignment(string alignmentStr) {
        if (Enum.TryParse<HorizontalAlignment>(alignmentStr, true, out var align)) {
            ActiveSheet?.ApplyStyleToSelection(cell => {
                cell.CellAlignment = align;
            });
        }
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