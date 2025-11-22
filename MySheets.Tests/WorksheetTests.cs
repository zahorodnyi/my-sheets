using Xunit;
using MySheets.Core.Models;

namespace MySheets.Tests;

public class WorksheetTests {
    [Fact]
    public void SetCell_ShouldStoreDataInSparseMatrix_AtHighCoordinates() {
        var worksheet = new Worksheet();
        int targetRow = 100;
        int targetCol = 50;
        string expectedValue = "Test Data";

        Assert.Equal(0, worksheet.ActiveCellCount);

        worksheet.SetCell(targetRow, targetCol, expectedValue);

        var retrievedCell = worksheet.GetCell(targetRow, targetCol);
        
        Assert.Equal(expectedValue, retrievedCell.Expression);
        Assert.Equal(targetRow, retrievedCell.Row);
        Assert.Equal(targetCol, retrievedCell.Column);
        Assert.Equal(1, worksheet.ActiveCellCount);
    }

    [Fact]
    public void Cell_ShouldAutoDetectType() {
        var cell = new Cell(0, 0);

        cell.Expression = "123.45";
        Assert.Equal(CellType.Number, cell.Type);

        cell.Expression = "=SUM(A1:A5)";
        Assert.Equal(CellType.Formula, cell.Type);

        cell.Expression = "Hello World";
        Assert.Equal(CellType.Text, cell.Type);
    }
}