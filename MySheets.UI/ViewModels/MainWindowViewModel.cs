using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using Avalonia.Layout; 
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.IO;
using MySheets.Core.Domain;

namespace MySheets.UI.ViewModels;

public partial class MainWindowViewModel : ObservableObject {
    private readonly FileService _fileService;
    private int _untitledCount = 1;

    public ObservableCollection<SheetEditor.SheetViewModel> Sheets { get; } = new();

    [ObservableProperty] 
    private SheetEditor.SheetViewModel? _activeSheet;

    public double CurrentFontSize {
        get => ActiveSheet?.SelectedCell?.FontSize ?? 12.0;
        set {
            if (ActiveSheet != null) {
                ActiveSheet.SetFontSizeForSelection(value);
                OnPropertyChanged();
            }
        }
    }

    public List<string> TextColors { get; } = new() {
        "#000000", "#434343", "#666666", "#999999", "#FFFFFF", "#CC0000", 
        "#E69138", "#F1C232", "#6AA84F", "#3D85C6", "#073763", "#674EA7"
    };

    public List<string> FillColors { get; } = new() {
        "#FFFFFF", "#F3F3F3", "#B7B7B7", "#F4CCCC", "#FCE5CD", "#FFF2CC", 
        "#D9EAD3", "#D0E0E3", "#CFE2F3", "#D9D2E9", "#EAD1DC"
    };

    [ObservableProperty]
    private string _currentTextColor = "#000000";

    [ObservableProperty]
    private string _currentFillColor = "Transparent";


    public MainWindowViewModel() {
        _fileService = new FileService();
        AddNewSheet();
    }

    protected override void OnPropertyChanged(PropertyChangedEventArgs e) {
        base.OnPropertyChanged(e);
        if (e.PropertyName == nameof(ActiveSheet)) {
            if (ActiveSheet != null) {
                ActiveSheet.PropertyChanged += ActiveSheet_PropertyChanged;
            }
            OnPropertyChanged(nameof(CurrentFontSize));
        }
    }

    private void ActiveSheet_PropertyChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(SheetEditor.SheetViewModel.SelectedCell)) {
            OnPropertyChanged(nameof(CurrentFontSize));
        }
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
        if (ActiveSheet != null) {
            double current = CurrentFontSize;
            if (current < 72) {
                CurrentFontSize = current + 1;
            }
        }
    }
    
    [RelayCommand]
    private void DecreaseFontSize() {
        if (ActiveSheet != null) {
            double current = CurrentFontSize;
            if (current > 6) {
                CurrentFontSize = current - 1;
            }
        }
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
    private void ApplyBorder(string type) {
        ActiveSheet?.ApplyBorderSelection(type);
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
        if (Enum.TryParse<Avalonia.Layout.HorizontalAlignment>(alignmentStr, true, out var align)) {
            ActiveSheet?.ApplyStyleToSelection(cell => {
                cell.CellAlignment = align;
            });
        }
    }

    [RelayCommand]
    public void SaveCurrentSheet() {
        if (ActiveSheet != null && !string.IsNullOrEmpty(ActiveSheet.FilePath)) {
            SaveData(ActiveSheet.FilePath);
        }
    }

    public void SaveData(string path) {
        if (ActiveSheet == null) return;
        
        var rowHeights = new Dictionary<int, double>();
        for (int i = 0; i < ActiveSheet.Rows.Count; i++) {
            rowHeights[i] = ActiveSheet.Rows[i].Height;
        }

        var colWidths = new Dictionary<int, double>();
        for (int i = 0; i < ActiveSheet.ColumnHeaders.Count; i++) {
            colWidths[i] = ActiveSheet.ColumnHeaders[i].Width;
        }

        _fileService.Save(path, ActiveSheet.Worksheet.Cells, rowHeights, colWidths);
        
        ActiveSheet.FilePath = path;
        ActiveSheet.Name = System.IO.Path.GetFileNameWithoutExtension(path);
    }

    public void LoadData(string path) {
        var sheetDto = _fileService.Load(path);
        
        var fileName = System.IO.Path.GetFileNameWithoutExtension(path);
        var newSheet = new SheetEditor.SheetViewModel(fileName);
        
        foreach (var cellDto in sheetDto.Cells) {
            var cell = newSheet.Worksheet.GetCell(cellDto.Row, cellDto.Col);
            
            newSheet.Worksheet.SetCell(cellDto.Row, cellDto.Col, cellDto.Expression);
            
            cell.FontSize = cellDto.FontSize;
            cell.IsBold = cellDto.IsBold;
            cell.IsItalic = cellDto.IsItalic;
            cell.TextColor = cellDto.TextColor;
            cell.BackgroundColor = cellDto.BackgroundColor;
            cell.BorderThickness = cellDto.BorderThickness;
            cell.TextAlignment = cellDto.TextAlignment;
            
            var cellVm = newSheet.Rows[cellDto.Row].Cells[cellDto.Col];
            cellVm.Refresh();
        }

        foreach (var kvp in sheetDto.RowHeights) {
            if (kvp.Key < newSheet.Rows.Count) {
                newSheet.Rows[kvp.Key].Height = kvp.Value;
                newSheet.Rows[kvp.Key].IsHeightManual = true;
            }
        }

        foreach (var kvp in sheetDto.ColumnWidths) {
            if (kvp.Key < newSheet.ColumnHeaders.Count) {
                newSheet.ColumnHeaders[kvp.Key].Width = kvp.Value;
            }
        }

        newSheet.FilePath = path;
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