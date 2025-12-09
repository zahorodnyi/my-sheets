using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text; 
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using MySheets.Core.Common;
using MySheets.Core.Domain;

namespace MySheets.UI.ViewModels.SheetEditor;

public class InternalClipboardData {
    public string TsvContent { get; set; } = string.Empty;
    public List<CellClipboardModel> Cells { get; set; } = new();
    public int RowCount { get; set; }
    public int ColCount { get; set; }
}

public class CellClipboardModel {
    public int RelRow { get; set; }
    public int RelCol { get; set; }
    public string Expression { get; set; } = string.Empty;
    public string BackgroundColor { get; set; } = "Transparent";
    public string TextColor { get; set; } = "#000000";
    public bool IsBold { get; set; }
    public bool IsItalic { get; set; }
    public double FontSize { get; set; }
    public string BorderThickness { get; set; } = "0,0,1,1";
    public string TextAlignment { get; set; } = "Left";
}

public partial class SheetViewModel : ObservableObject {
    private static InternalClipboardData? _internalClipboard;

    public Worksheet Worksheet { get; }
    public UndoRedoManager History { get; }

    private int _anchorRow;
    private int _anchorCol;
    private int _currentRow;
    private int _currentCol;

    [ObservableProperty] private string _name;
    
    [ObservableProperty] private double _selectionX;
    [ObservableProperty] private double _selectionY;
    [ObservableProperty] private double _selectionWidth;
    [ObservableProperty] private double _selectionHeight;
    [ObservableProperty] private bool _isSelectionVisible;

    [ObservableProperty] private double _refSelectionX;
    [ObservableProperty] private double _refSelectionY;
    [ObservableProperty] private double _refSelectionWidth;
    [ObservableProperty] private double _refSelectionHeight;
    [ObservableProperty] private bool _isRefSelectionVisible;

    [ObservableProperty] private CellViewModel? _selectedCell;
    [ObservableProperty] private string _currentAddress = "A1";

    private RowViewModel? _activeRowHeader;
    private ColumnViewModel? _activeColHeader;

    public ObservableCollection<ColumnViewModel> ColumnHeaders { get; }
    public ObservableCollection<RowViewModel> Rows { get; }

    public SheetViewModel(string name) {
        Name = name;
        Worksheet = new Worksheet();
        History = new UndoRedoManager(10); 
        
        ColumnHeaders = new ObservableCollection<ColumnViewModel>();
        Rows = new ObservableCollection<RowViewModel>();

        const int RowCount = 50;
        const int ColCount = 25;

        for (int c = 0; c < ColCount; c++) {
            ColumnHeaders.Add(new ColumnViewModel(MainWindowViewModel.GetColumnName(c)));
        }

        for (var r = 0; r < RowCount; r++) {
            var rowCells = new List<CellViewModel>();
            for (var c = 0; c < ColCount; c++) {
                var cellModel = Worksheet.GetCell(r, c);
                rowCells.Add(new CellViewModel(cellModel, Worksheet, ColumnHeaders[c], this));
            }
            Rows.Add(new RowViewModel(rowCells, r + 1));
        }

        Worksheet.CellStateChanged += OnCellStateChanged;
        History.StateChanged += (s, e) => {
            UndoCommand.NotifyCanExecuteChanged();
            RedoCommand.NotifyCanExecuteChanged();
        };
    }

    [RelayCommand(CanExecute = nameof(CanUndo))]
    public void Undo() => History.Undo();

    [RelayCommand(CanExecute = nameof(CanRedo))]
    public void Redo() => History.Redo();

    private bool CanUndo() => History.CanUndo;
    private bool CanRedo() => History.CanRedo;

    private void OnCellStateChanged(int row, int col) {
        if (row < Rows.Count && col < Rows[row].Cells.Count) {
             var vm = Rows[row].Cells[col];
             vm.Refresh();
        }
    }

    public void StartSelection(int row, int col) {
        _anchorRow = row;
        _anchorCol = col;
        
        if (row >= 0 && row < Rows.Count && col >= 0 && col < Rows[row].Cells.Count) {
            SelectedCell = Rows[row].Cells[col];
        }
        
        if (_activeRowHeader != null) _activeRowHeader.IsActive = false;
        if (_activeColHeader != null) _activeColHeader.IsActive = false;

        if (row >= 0 && row < Rows.Count) {
            _activeRowHeader = Rows[row];
            _activeRowHeader.IsActive = true;
        }
        if (col >= 0 && col < ColumnHeaders.Count) {
            _activeColHeader = ColumnHeaders[col];
            _activeColHeader.IsActive = true;
        }

        IsSelectionVisible = true;
        UpdateSelection(row, col);
    }

    public void UpdateSelection(int row, int col) {
        if (row < 0 || col < 0) return;
        _currentRow = row;
        _currentCol = col;
        UpdateSelectionGeometry(row, col);
        UpdateAddressText();
    }

    public void ApplyStyleToSelection(Action<CellViewModel> applyAction) {
        int r1 = Math.Min(_anchorRow, _currentRow);
        int r2 = Math.Max(_anchorRow, _currentRow);
        int c1 = Math.Min(_anchorCol, _currentCol);
        int c2 = Math.Max(_anchorCol, _currentCol);

        History.StartGroup();

        try {
            for (int r = r1; r <= r2; r++) {
                if (r >= Rows.Count) continue;
                var row = Rows[r];
                for (int c = c1; c <= c2; c++) {
                    if (c >= row.Cells.Count) continue;
                    applyAction(row.Cells[c]);
                }
            }
        }
        finally {
            History.EndGroup();
        }
    }

    public void SetFontSizeForSelection(double newSize) {
        int r1 = Math.Min(_anchorRow, _currentRow);
        int r2 = Math.Max(_anchorRow, _currentRow);
        int c1 = Math.Min(_anchorCol, _currentCol);
        int c2 = Math.Max(_anchorCol, _currentCol);

        double requiredHeight = newSize * 1.4;

        History.StartGroup();
        try {
            for (int r = r1; r <= r2; r++) {
                if (r >= Rows.Count) continue;
                
                var row = Rows[r];

                if (row.IsHeightManual) {
                    if (row.Height < requiredHeight) {
                        row.Height = requiredHeight;
                    }
                } 
                else {
                    row.Height = Math.Max(25, requiredHeight);
                }

                for (int c = c1; c <= c2; c++) {
                    if (c >= row.Cells.Count) continue;
                    row.Cells[c].FontSize = newSize;
                }
            }
        }
        finally {
            History.EndGroup();
        }
    }
    
    public void ApplyBorderSelection(string mode) {
        int r1 = Math.Min(_anchorRow, _currentRow);
        int r2 = Math.Max(_anchorRow, _currentRow);
        int c1 = Math.Min(_anchorCol, _currentCol);
        int c2 = Math.Max(_anchorCol, _currentCol);

        History.StartGroup();
        try {
            for (int r = r1; r <= r2; r++) {
                if (r >= Rows.Count) continue;
                var row = Rows[r];
                
                for (int c = c1; c <= c2; c++) {
                    if (c >= row.Cells.Count) continue;
                    var cell = row.Cells[c];

                    if (mode == "None") {
                        cell.SetBorder("0,0,1,1");
                    }
                    else if (mode == "All") {
                        int left = (c == c1) ? 1 : 0;
                        int top = (r == r1) ? 1 : 0;
                        
                        cell.SetBorder($"{left},{top},1.00,1.00");
                    }
                    else if (mode == "Outside") {
                        bool isLeftEdge = (c == c1);
                        bool isTopEdge = (r == r1);
                        bool isRightEdge = (c == c2);
                        bool isBottomEdge = (r == r2);

                        if (!isLeftEdge && !isTopEdge && !isRightEdge && !isBottomEdge) continue;

                        var (curL, curT, curR, curB) = GetCurrentBorderValues(cell);

                        string newL = isLeftEdge ? "1" : (curL > 0 ? "1.00" : "0");
                        string newT = isTopEdge ? "1" : (curT > 0 ? "1.00" : "0");
                        string newR = isRightEdge ? "1.00" : (curR > 0 ? "1.00" : "0"); 
                        string newB = isBottomEdge ? "1.00" : (curB > 0 ? "1.00" : "0");
                        
                        cell.SetBorder($"{newL},{newT},{newR},{newB}");
                    }
                }
            }
        }
        finally {
            History.EndGroup();
        }
    }
    
    public void ClearSelectionContent() {
        int r1 = Math.Min(_anchorRow, _currentRow);
        int r2 = Math.Max(_anchorRow, _currentRow);
        int c1 = Math.Min(_anchorCol, _currentCol);
        int c2 = Math.Max(_anchorCol, _currentCol);

        History.StartGroup();
        try {
            for (int r = r1; r <= r2; r++) {
                if (r >= Rows.Count) continue;
                var row = Rows[r];
                for (int c = c1; c <= c2; c++) {
                    if (c >= row.Cells.Count) continue;
                    
                    row.Cells[c].Expression = string.Empty;
                }
            }
        }
        finally {
            History.EndGroup();
        }
    }

    private (int, int, int, int) GetCurrentBorderValues(CellViewModel vm) {
        var parts = vm.BorderThickness.ToString().Split(',');
        if (parts.Length == 4) {
            return ((int)vm.BorderThickness.Left, (int)vm.BorderThickness.Top, (int)vm.BorderThickness.Right, (int)vm.BorderThickness.Bottom);
        }
        return (0, 0, 1, 1);
    }

    private void UpdateAddressText() {
        string startCol = CellReferenceUtility.GetColumnName(_anchorCol);
        string startAddr = $"{startCol}{_anchorRow + 1}";

        if (_anchorRow == _currentRow && _anchorCol == _currentCol) {
            CurrentAddress = startAddr;
        } 
        else {
            int r1 = Math.Min(_anchorRow, _currentRow);
            int r2 = Math.Max(_anchorRow, _currentRow);
            int c1 = Math.Min(_anchorCol, _currentCol);
            int c2 = Math.Max(_anchorCol, _currentCol);

            string tlCol = CellReferenceUtility.GetColumnName(c1);
            string brCol = CellReferenceUtility.GetColumnName(c2);
            CurrentAddress = $"{tlCol}{r1 + 1}:{brCol}{r2 + 1}";
        }
    }

    private void UpdateSelectionGeometry(int currentRow, int currentCol) {
        int startR = Math.Min(_anchorRow, currentRow);
        int endR = Math.Max(_anchorRow, currentRow);
        int startC = Math.Min(_anchorCol, currentCol);
        int endC = Math.Max(_anchorCol, currentCol);

        double x = 0;
        double y = 0;
        double w = 0;
        double h = 0;

        for (int c = 0; c < startC; c++) x += ColumnHeaders[c].Width;
        for (int r = 0; r < startR; r++) y += Rows[r].Height;
        for (int c = startC; c <= endC; c++) w += ColumnHeaders[c].Width;
        for (int r = startR; r <= endR; r++) h += Rows[r].Height;

        SelectionX = x;
        SelectionY = y;
        SelectionWidth = w;
        SelectionHeight = h;
    }

    public void ShowRefSelection(int row, int col, int endRow = -1, int endCol = -1) {
         if (endRow == -1) endRow = row;
         if (endCol == -1) endCol = col;
         
         int r1 = Math.Min(row, endRow);
         int r2 = Math.Max(row, endRow);
         int c1 = Math.Min(col, endCol);
         int c2 = Math.Max(col, endCol);

         double x = 0, y = 0, w = 0, h = 0;
         
         for (int c = 0; c < c1; c++) x += ColumnHeaders[c].Width;
         for (int r = 0; r < r1; r++) y += Rows[r].Height;
         for (int c = c1; c <= c2; c++) w += ColumnHeaders[c].Width;
         for (int r = r1; r <= r2; r++) h += Rows[r].Height;

         RefSelectionX = x;
         RefSelectionY = y;
         RefSelectionWidth = w;
         RefSelectionHeight = h;
         IsRefSelectionVisible = true;
    }

    public void HideRefSelection() {
        IsRefSelectionVisible = false;
    }

    [RelayCommand]
    public void AddRow() {
        History.StartGroup();
        int rowIndex = Rows.Count;
        var rowCells = new List<CellViewModel>();
        for (int c = 0; c < ColumnHeaders.Count; c++) {
            var cellModel = Worksheet.GetCell(rowIndex, c);
            rowCells.Add(new CellViewModel(cellModel, Worksheet, ColumnHeaders[c], this));
        }
        Rows.Add(new RowViewModel(rowCells, rowIndex + 1));
        History.EndGroup();
    }

    [RelayCommand]
    public void RemoveRow() {
        if (Rows.Count > 0) Rows.RemoveAt(Rows.Count - 1);
    }

    [RelayCommand]
    public void AddColumn() {
        History.StartGroup();
        int colIndex = ColumnHeaders.Count;
        var newColumn = new ColumnViewModel(MainWindowViewModel.GetColumnName(colIndex));
        ColumnHeaders.Add(newColumn);
        for (int r = 0; r < Rows.Count; r++) {
            var cellModel = Worksheet.GetCell(r, colIndex);
            Rows[r].Cells.Add(new CellViewModel(cellModel, Worksheet, newColumn, this));
        }
        History.EndGroup();
    }

    [RelayCommand]
    public void RemoveColumn() {
        if (ColumnHeaders.Count > 0) {
            int lastColIndex = ColumnHeaders.Count - 1;
            ColumnHeaders.RemoveAt(lastColIndex);
            foreach (var row in Rows) {
                if (row.Cells.Count > 0) row.Cells.RemoveAt(row.Cells.Count - 1);
            }
        }
    }
    
    public bool ValidateFormula(string formula) {
        return Worksheet.IsFormulaValid(formula);
    }

    public string CopySelection() {
        int r1 = Math.Min(_anchorRow, _currentRow);
        int r2 = Math.Max(_anchorRow, _currentRow);
        int c1 = Math.Min(_anchorCol, _currentCol);
        int c2 = Math.Max(_anchorCol, _currentCol);

        var sb = new StringBuilder();
        var clipboardData = new InternalClipboardData {
            RowCount = r2 - r1 + 1,
            ColCount = c2 - c1 + 1
        };

        for (int r = r1; r <= r2; r++) {
            if (r >= Rows.Count) continue;
            var row = Rows[r];
            
            for (int c = c1; c <= c2; c++) {
                if (c >= row.Cells.Count) continue;
                
                var cellVm = row.Cells[c];
                
                string val = cellVm.Expression;
                val = val.Replace("\t", " ").Replace("\n", " "); 
                sb.Append(val);
                if (c < c2) sb.Append("\t");

                clipboardData.Cells.Add(new CellClipboardModel {
                    RelRow = r - r1,
                    RelCol = c - c1,
                    Expression = cellVm.Expression,
                    BackgroundColor = cellVm.Background.ToString() ?? "Transparent",
                    TextColor = cellVm.Foreground.ToString() ?? "#000000",
                    IsBold = cellVm.IsBold,
                    IsItalic = cellVm.IsItalic,
                    FontSize = cellVm.FontSize,
                    BorderThickness = cellVm.BorderThickness.ToString(),
                    TextAlignment = cellVm.CellAlignment.ToString()
                });
            }
            if (r < r2) sb.Append("\r\n");
        }

        string tsv = sb.ToString();
        clipboardData.TsvContent = tsv;
        
        _internalClipboard = clipboardData;
        
        return tsv;
    }

    public void PasteData(string systemClipboardText) {
        if (string.IsNullOrEmpty(systemClipboardText)) return;

        History.StartGroup();
        try {
            int startRow = _anchorRow;
            int startCol = _anchorCol;

            bool useInternal = _internalClipboard != null && _internalClipboard.TsvContent == systemClipboardText;

            if (useInternal && _internalClipboard != null) {
                foreach (var item in _internalClipboard.Cells) {
                    int targetRow = startRow + item.RelRow;
                    int targetCol = startCol + item.RelCol;

                    if (targetRow < Rows.Count && targetCol < ColumnHeaders.Count) {
                        var cell = Rows[targetRow].Cells[targetCol];
                        cell.Expression = item.Expression;
                        cell.SetBackgroundColor(item.BackgroundColor);
                        cell.SetTextColor(item.TextColor);
                        cell.IsBold = item.IsBold;
                        cell.IsItalic = item.IsItalic;
                        cell.FontSize = item.FontSize;
                        cell.SetBorder(item.BorderThickness);
                        if (Enum.TryParse<Avalonia.Layout.HorizontalAlignment>(item.TextAlignment, out var align)) {
                            cell.CellAlignment = align;
                        }
                    }
                }
                
                UpdateSelection(
                    Math.Min(Rows.Count - 1, startRow + _internalClipboard.RowCount - 1),
                    Math.Min(ColumnHeaders.Count - 1, startCol + _internalClipboard.ColCount - 1)
                );
            }
            else {
                var rows = systemClipboardText.Replace("\r\n", "\n").Split('\n');
                for (int i = 0; i < rows.Length; i++) {
                    var rowData = rows[i];
                    if (i == rows.Length - 1 && string.IsNullOrEmpty(rowData)) continue; 

                    var cols = rowData.Split('\t');
                    int targetRow = startRow + i;

                    if (targetRow >= Rows.Count) break;

                    for (int j = 0; j < cols.Length; j++) {
                        int targetCol = startCol + j;
                        if (targetCol >= ColumnHeaders.Count) break;

                        var cell = Rows[targetRow].Cells[targetCol];
                        cell.Expression = cols[j].Trim();
                    }
                }
                
                int endRow = Math.Min(Rows.Count - 1, startRow + rows.Length - (string.IsNullOrEmpty(rows[^1]) ? 2 : 1));
                int maxCols = 0;
                foreach(var r in rows) maxCols = Math.Max(maxCols, r.Split('\t').Length);
                int endCol = Math.Min(ColumnHeaders.Count - 1, startCol + maxCols - 1);

                UpdateSelection(endRow, endCol);
            }
        }
        finally {
            History.EndGroup();
        }
    }
}