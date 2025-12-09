using System.Text.Json;
using System.Text.Json.Serialization;
using ClosedXML.Excel;
using MySheets.Core.Domain;

namespace MySheets.Core.IO;

public record CellDto(
    int Row, 
    int Col, 
    string Expression,
    double FontSize,
    bool IsBold,
    bool IsItalic,
    string TextColor,
    string BackgroundColor,
    string BorderThickness,
    string TextAlignment
);

public class SheetDto {
    public List<CellDto> Cells { get; set; } = new();
    public Dictionary<int, double> RowHeights { get; set; } = new();
    public Dictionary<int, double> ColumnWidths { get; set; } = new();
}

public class FileService {
    private const double DefaultFontSize = 12.0;
    private const string DefaultTextColor = "#000000";
    private const string DefaultBackgroundColor = "Transparent";
    private const string DefaultBorderThickness = "0,0,1,1";
    private const string DefaultTextAlignment = "Left";

    private const double PxToPoints = 0.75; 
    private const double PxToCharWidth = 1.0 / 7.0; 

    public void Save(string path, IEnumerable<Cell> cells, Dictionary<int, double> rowHeights, Dictionary<int, double> colWidths) {
        if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) {
            SaveAsXlsx(path, cells, rowHeights, colWidths);
        } 
        else {
            var cellDtos = cells
                .Where(c => 
                    !string.IsNullOrEmpty(c.Expression) ||
                    c.FontSize != DefaultFontSize ||
                    c.IsBold ||
                    c.IsItalic ||
                    c.TextColor != DefaultTextColor ||
                    c.BackgroundColor != DefaultBackgroundColor ||
                    c.BorderThickness != DefaultBorderThickness ||
                    c.TextAlignment != DefaultTextAlignment
                )
                .Select(c => new CellDto(
                    c.Row, 
                    c.Col, 
                    c.Expression,
                    c.FontSize,
                    c.IsBold,
                    c.IsItalic,
                    c.TextColor,
                    c.BackgroundColor,
                    c.BorderThickness,
                    c.TextAlignment
                )).ToList();

            var sheetData = new SheetDto {
                Cells = cellDtos,
                RowHeights = rowHeights,
                ColumnWidths = colWidths
            };
                
            var options = new JsonSerializerOptions { 
                WriteIndented = true,
                DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull 
            };
            var json = JsonSerializer.Serialize(sheetData, options);
            File.WriteAllText(path, json);
        }
    }

    private void SaveAsXlsx(string path, IEnumerable<Cell> cells, Dictionary<int, double> rowHeights, Dictionary<int, double> colWidths) {
        XLWorkbook workbook;
        bool isExisting = File.Exists(path);
        
        try {
            if (isExisting) {
                workbook = new XLWorkbook(path);
            } else {
                workbook = new XLWorkbook();
            }
        } catch {
            workbook = new XLWorkbook();
        }

        using (workbook) {
            var worksheetName = "Sheet1";
            IXLWorksheet worksheet;

            if (workbook.Worksheets.TryGetWorksheet(worksheetName, out var existingSheet)) {
                worksheet = existingSheet;
                worksheet.Clear(); 
            } else {
                worksheet = workbook.Worksheets.Add(worksheetName);
            }

            foreach (var kvp in rowHeights) {
                worksheet.Row(kvp.Key + 1).Height = kvp.Value * PxToPoints;
            }

            foreach (var kvp in colWidths) {
                worksheet.Column(kvp.Key + 1).Width = kvp.Value * PxToCharWidth;
            }

            foreach (var cell in cells) {
                bool hasContent = !string.IsNullOrEmpty(cell.Expression);
                bool hasStyle = cell.BackgroundColor != DefaultBackgroundColor || 
                                cell.BorderThickness != DefaultBorderThickness ||
                                cell.IsBold || cell.IsItalic || cell.FontSize != DefaultFontSize;

                if (!hasContent && !hasStyle) continue;

                var xlCell = worksheet.Cell(cell.Row + 1, cell.Col + 1);

                if (cell.Type == CellType.Formula) {
                    string formula = cell.Expression.StartsWith("=") ? cell.Expression.Substring(1) : cell.Expression;
                    xlCell.FormulaA1 = formula;
                } 
                else {
                    xlCell.Value = cell.Expression;
                }

                if (cell.IsBold) xlCell.Style.Font.Bold = true;
                if (cell.IsItalic) xlCell.Style.Font.Italic = true;
                if (cell.FontSize > 0) xlCell.Style.Font.FontSize = cell.FontSize;
                
                if (!string.IsNullOrEmpty(cell.TextColor) && cell.TextColor != DefaultTextColor) {
                    try {
                        string colorHex = FixHexColor(cell.TextColor);
                        xlCell.Style.Font.FontColor = XLColor.FromHtml(colorHex);
                    } 
                    catch {  }
                }
                
                if (!string.IsNullOrEmpty(cell.BackgroundColor) && cell.BackgroundColor != DefaultBackgroundColor) {
                    try {
                        string colorHex = FixHexColor(cell.BackgroundColor);
                        xlCell.Style.Fill.BackgroundColor = XLColor.FromHtml(colorHex);
                    } 
                    catch {  }
                }
                
                xlCell.Style.Alignment.Horizontal = cell.TextAlignment switch {
                    "Center" => XLAlignmentHorizontalValues.Center,
                    "Right" => XLAlignmentHorizontalValues.Right,
                    _ => XLAlignmentHorizontalValues.Left
                };
                
                if (cell.BorderThickness != DefaultBorderThickness && !string.IsNullOrEmpty(cell.BorderThickness)) {
                    var parts = cell.BorderThickness.Split(',');
                    
                    if (parts.Length == 4 && double.TryParse(parts[0], out double l) && double.TryParse(parts[1], out double t) && double.TryParse(parts[2], out double r) && double.TryParse(parts[3], out double b)) {
                        if (l > 0) { xlCell.Style.Border.LeftBorder = XLBorderStyleValues.Thin; xlCell.Style.Border.LeftBorderColor = XLColor.Black; }
                        if (t > 0) { xlCell.Style.Border.TopBorder = XLBorderStyleValues.Thin; xlCell.Style.Border.TopBorderColor = XLColor.Black; }
                        if (r > 0) { xlCell.Style.Border.RightBorder = XLBorderStyleValues.Thin; xlCell.Style.Border.RightBorderColor = XLColor.Black; }
                        if (b > 0) { xlCell.Style.Border.BottomBorder = XLBorderStyleValues.Thin; xlCell.Style.Border.BottomBorderColor = XLColor.Black; }
                    }
                }
            }
            
            workbook.SaveAs(path);
        }
    }

    private string FixHexColor(string hex) {
        if (hex.Length == 9 && hex.StartsWith("#")) {
            return "#" + hex.Substring(3);
        }
        return hex;
    }

    public SheetDto Load(string path) {
        if (!File.Exists(path)) throw new FileNotFoundException("File not found", path);

        if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) {
            return LoadFromXlsx(path);
        }

        var json = File.ReadAllText(path);
        try {
            var sheetDto = JsonSerializer.Deserialize<SheetDto>(json);
            if (sheetDto != null) return sheetDto;
        } 
        catch { }

        var oldCells = JsonSerializer.Deserialize<IEnumerable<CellDto>>(json) ?? new List<CellDto>();
        return new SheetDto { Cells = oldCells.ToList() };
    }

    private SheetDto LoadFromXlsx(string path) {
        var result = new SheetDto();
        using var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheets.FirstOrDefault() ?? workbook.Worksheet(1);

        var usedRows = worksheet.RowsUsed();
        foreach(var row in usedRows) {
            result.RowHeights[row.RowNumber() - 1] = row.Height / PxToPoints;
        }

        var usedCols = worksheet.ColumnsUsed();
        foreach(var col in usedCols) {
            result.ColumnWidths[col.ColumnNumber() - 1] = col.Width / PxToCharWidth;
        }

        var range = worksheet.RangeUsed();
        if (range != null) {
            foreach (var cell in range.Cells()) {
                int row = cell.Address.RowNumber - 1; 
                int col = cell.Address.ColumnNumber - 1;

                string expression;
                
                if (cell.HasFormula) {
                    expression = "=" + cell.FormulaA1;
                } 
                else {
                    expression = cell.Value.ToString();
                }

                var style = cell.Style;

                string bgColor = DefaultBackgroundColor;
                if (style.Fill.BackgroundColor != XLColor.NoColor) {
                    try {
                        var c = style.Fill.BackgroundColor.Color;
                        bgColor = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                    }
                    catch { }
                }
                
                string textColor = DefaultTextColor;
                if (style.Font.FontColor != XLColor.NoColor) {
                    try {
                        var c = style.Font.FontColor.Color;
                        textColor = $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                    }
                    catch { }
                }

                double fontSize = style.Font.FontSize;
                bool isBold = style.Font.Bold;
                bool isItalic = style.Font.Italic;
                string align = style.Alignment.Horizontal.ToString();

                double bLeft = style.Border.LeftBorder != XLBorderStyleValues.None ? 1 : 0;
                double bTop = style.Border.TopBorder != XLBorderStyleValues.None ? 1 : 0;
                double bRight = style.Border.RightBorder != XLBorderStyleValues.None ? 1 : 0;
                double bBottom = style.Border.BottomBorder != XLBorderStyleValues.None ? 1 : 0;
                
                string borderThickness = $"{bLeft},{bTop},{bRight},{bBottom}";

                if (string.IsNullOrEmpty(expression) && 
                    bgColor == DefaultBackgroundColor && 
                    borderThickness == "0,0,0,0") { 
                    continue;
                }

                result.Cells.Add(new CellDto(
                    row, 
                    col, 
                    expression,
                    fontSize,
                    isBold,
                    isItalic,
                    textColor,
                    bgColor,
                    borderThickness,
                    align
                ));
            }
        }
        
        return result;
    }
}