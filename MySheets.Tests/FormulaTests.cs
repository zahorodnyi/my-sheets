using MySheets.Core.Models;
using Xunit;

namespace MySheets.Tests;

public class FormulaTests {
    [Fact]
    public void SetCell_ShouldCalculateSimpleAddition() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=2+2");
        var cell = sheet.GetCell(0, 0);

        Assert.Equal(CellType.Formula, cell.Type);
        Assert.Equal(4.0, cell.Value);
    }

    [Fact]
    public void SetCell_ShouldCalculateOrderOfOperations() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=2+3*4"); 
        var cell = sheet.GetCell(0, 0);

        Assert.Equal(14.0, cell.Value);
    }

    [Fact]
    public void SetCell_ShouldCalculateDivisionAndSubtraction() {
        var sheet = new Worksheet();
        sheet.SetCell(1, 1, "=10/2-1");
        var cell = sheet.GetCell(1, 1);

        Assert.Equal(4.0, cell.Value);
    }

    [Fact]
    public void SetCell_ShouldHandleDecimals() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=2.5*2");
        var cell = sheet.GetCell(0, 0);

        Assert.Equal(5.0, cell.Value);
    }

    [Fact]
    public void SetCell_ShouldReturnErrorOnInvalidSyntax() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=2++2"); 
        var cell = sheet.GetCell(0, 0);

        Assert.Equal("#ERROR!", cell.Value);
    }
    
    [Fact]
    public void SetCell_RegularText_ShouldNotBeEvaluated() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "2+2"); 
        var cell = sheet.GetCell(0, 0);

        Assert.Equal(CellType.Text, cell.Type);
        Assert.Equal("2+2", cell.Value);
    }

    [Fact]
    public void SetCell_ShouldPrioritizeParentheses() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=(2+3)*4");
        var cell = sheet.GetCell(0, 0);

        Assert.Equal(20.0, cell.Value);
    }

    [Fact]
    public void SetCell_ShouldHandleNestedParentheses() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=((2+2)*3)/4");
        var cell = sheet.GetCell(0, 0);

        Assert.Equal(3.0, cell.Value);
    }

    [Fact]
    public void SetCell_ShouldEvaluateLeftToRight() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=10-3-2");
        var cell = sheet.GetCell(0, 0);

        Assert.Equal(5.0, cell.Value);
    }

    [Fact]
    public void SetCell_ShouldIgnoreWhitespace() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=  4  * 5  ");
        var cell = sheet.GetCell(0, 0);

        Assert.Equal(20.0, cell.Value);
    }
    
    [Fact]
    public void SetCell_ShouldUpdateDependentCells() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "10");    
        sheet.SetCell(0, 1, "=A1*2"); 
        
        sheet.SetCell(0, 0, "5");
        
        var b1 = sheet.GetCell(0, 1);
        Assert.Equal(10.0, b1.Value); 
    }

    [Fact]
    public void SetCell_ShouldUpdateChainedDependencies() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "2");      
        sheet.SetCell(0, 1, "=A1*2"); 
        sheet.SetCell(0, 2, "=B1+1");  
        
        sheet.SetCell(0, 0, "5");
        
        var c1 = sheet.GetCell(0, 2);
        Assert.Equal(11.0, c1.Value);
    }
    
}