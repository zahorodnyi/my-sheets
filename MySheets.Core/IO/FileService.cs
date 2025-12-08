using System.Text.Json;
using ClosedXML.Excel;
using MySheets.Core.Domain;

namespace MySheets.Core.IO;

public record CellDto(int Row, int Col, string Expression);

public class FileService {
    public void Save(string path, IEnumerable<Cell> cells) {
        if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) {
            SaveAsXlsx(path, cells);
        } 
        else {
            var data = cells.Select(c => new CellDto(c.Row, c.Col, c.Expression));
            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(data, options);
            File.WriteAllText(path, json);
        }
    }

    private void SaveAsXlsx(string path, IEnumerable<Cell> cells) {
        using var workbook = new XLWorkbook();
        var worksheet = workbook.Worksheets.Add("Sheet1");

        foreach (var cell in cells) {
            var xlCell = worksheet.Cell(cell.Row + 1, cell.Col + 1);

            if (cell.Type == CellType.Formula) {
                string formula = cell.Expression.StartsWith("=") ? cell.Expression.Substring(1) : cell.Expression;
                xlCell.FormulaA1 = formula;
            } 
            else {
                if (cell.Value is double num) {
                    xlCell.Value = num;
                } 
                else {
                    xlCell.Value = cell.Expression;
                }
            }

            if (cell.IsBold) xlCell.Style.Font.Bold = true;
            if (cell.IsItalic) xlCell.Style.Font.Italic = true;
            if (cell.FontSize > 0) xlCell.Style.Font.FontSize = cell.FontSize;
            
            if (!string.IsNullOrEmpty(cell.TextColor)) {
                try {
                    xlCell.Style.Font.FontColor = XLColor.FromHtml(cell.TextColor);
                } 
                catch {  }
            }
            
            if (!string.IsNullOrEmpty(cell.BackgroundColor) && cell.BackgroundColor != "Transparent") {
                try {
                    xlCell.Style.Fill.BackgroundColor = XLColor.FromHtml(cell.BackgroundColor);
                } 
                catch {  }
            }
            
            if (!string.IsNullOrEmpty(cell.BorderThickness) && cell.BorderThickness.Contains("1")) {
                 xlCell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                 xlCell.Style.Border.OutsideBorderColor = XLColor.Black;
            }
        }
        
        worksheet.Columns().AdjustToContents();
        workbook.SaveAs(path);
    }

    public IEnumerable<CellDto> Load(string path) {
        if (!File.Exists(path)) throw new FileNotFoundException("File not found", path);

        if (path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)) {
            return LoadFromXlsx(path);
        }

        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<IEnumerable<CellDto>>(json) ?? new List<CellDto>();
    }

    private IEnumerable<CellDto> LoadFromXlsx(string path) {
        var result = new List<CellDto>();
        using var workbook = new XLWorkbook(path);
        var worksheet = workbook.Worksheet(1);

        foreach (var cell in worksheet.CellsUsed()) {
            int row = cell.Address.RowNumber - 1; 
            int col = cell.Address.ColumnNumber - 1;

            string expression;
            
            if (cell.HasFormula) {
                expression = "=" + cell.FormulaA1;
            } 
            else {
                expression = cell.Value.ToString();
            }

            result.Add(new CellDto(row, col, expression));
        }
        
        return result;
    }
}