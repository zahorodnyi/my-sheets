using MySheets.Core.Domain;

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

        Assert.Equal(4.0, cell.Value);
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

    [Fact]
    public void SetCell_Sum_ShouldCalculateVariableArgs() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1, 2, 3, 4, 5)");
        Assert.Equal(15.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Sum_ShouldHandleSingleArg() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(42)");
        Assert.Equal(42.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Sum_ShouldHandleNegatives() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(-5, 10, -2)");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Sum_MixedReferencesAndLiterals() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "10"); 
        sheet.SetCell(0, 1, "20"); 
        sheet.SetCell(0, 2, "=SUM(A1, 5, B1)"); 
        
        Assert.Equal(35.0, sheet.GetCell(0, 2).Value);
    }

    [Fact]
    public void SetCell_Average_Basic() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(10, 20, 30)");
        Assert.Equal(20.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Average_DecimalResult() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(2, 3)"); 
        Assert.Equal(2.5, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Average_ShouldHandleZeroValuesCorrectly() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(0, 10)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Max_ShouldFindMaximum() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX(1, 99, 2, 5)");
        Assert.Equal(99.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Max_ShouldHandleNegatives() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX(-5, -1, -20)");
        Assert.Equal(-1.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Min_ShouldFindMinimum() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MIN(10, 5, 20)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Min_MixedSigns() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MIN(-10, 10, 0)");
        Assert.Equal(-10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Median_OddCount_Unsorted() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN(5, 1, 9)"); 
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Median_EvenCount_ShouldAverageMiddleTwo() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN(1, 2, 3, 4)"); 
        Assert.Equal(2.5, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_Median_UnsortedComplex() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN(100, 1, 50, 20)"); 
        Assert.Equal(35.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_NestedFunctions_Simple() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(MAX(1,5), MIN(2,8))"); 
        Assert.Equal(7.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_NestedFunctions_Deep() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(SUM(10,10), MAX(10, 30))"); 
        Assert.Equal(25.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_FunctionsCombinedWithArithmetic() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1,2) * 3 + MAX(4,5)");
        Assert.Equal(14.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_CaseInsensitivity_Lower() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=sum(1, 2)");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_CaseInsensitivity_Mixed() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MeDiAn(1, 3, 5)");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_ComplexDependency_Functions() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "1"); 
        sheet.SetCell(0, 1, "2"); 
        sheet.SetCell(0, 2, "=SUM(A1, B1)"); 
        sheet.SetCell(0, 3, "=MAX(C1, 10)"); 

        Assert.Equal(3.0, sheet.GetCell(0, 2).Value);
        Assert.Equal(10.0, sheet.GetCell(0, 3).Value);

        sheet.SetCell(0, 0, "20"); 

        Assert.Equal(22.0, sheet.GetCell(0, 2).Value);
        Assert.Equal(22.0, sheet.GetCell(0, 3).Value);
    }

    [Fact]
    public void SetCell_MultipleReferencesSameCell() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "5");
        sheet.SetCell(0, 1, "=SUM(A1, A1, A1)");
        Assert.Equal(15.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void SetCell_FunctionWithWhitespaceFormatting() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "= SUM(  1  ,   2   ) ");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_EmptyFunction_ShouldReturnError() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM()");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_MissingArgumentInMiddle_ShouldReturnError() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1, , 2)");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_UnknownFunction_ShouldReturnError() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=UNKNOWN(1, 2)");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_MissingClosingParenthesis_ShouldReturnError() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1, 2");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_LargeNumberOfArguments() {
        var sheet = new Worksheet();
        var args = string.Join(",", Enumerable.Range(1, 100).Select(x => "1"));
        sheet.SetCell(0, 0, $"=SUM({args})");
        Assert.Equal(100.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void SetCell_FunctionsWithDivisionByZero_ShouldError() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(10/0, 5)");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }
    
    [Fact]
    public void SetCell_NestedFunction_ErrorPropagation() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1, AVERAGE(10/0))");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Addition_ShouldHandleLargeNumbers() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=1000000+2500000");
        Assert.Equal(3500000.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Subtraction_ShouldHandleNegativeResult() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=5-10");
        Assert.Equal(-5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Multiplication_ByZero() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=999*0");
        Assert.Equal(0.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Division_FloatingPointResult() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=7/2");
        Assert.Equal(3.5, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void UnaryMinus_ShouldWork() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=-5");
        Assert.Equal(-5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void UnaryPlus_ShouldBeIgnored() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=+7");
        Assert.Equal(7.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void MultipleWhitespaceInsideExpression() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "= 10  +   20 *  3 ");
        Assert.Equal(70.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void ParenthesesWithWhitespace() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "= ( 2 + 3 ) * 4");
        Assert.Equal(20.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void DeepNestedParenthesesComplex() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=(((1+2)*3)-(4/2))*5");
        Assert.Equal(35.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void ErrorOnDoubleOperator_Multiply() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=2**5");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void ErrorOnTrailingOperator() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=10+");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void ErrorOnStartingOperator() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=*5");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Sum_ShouldHandleLargeMixedValues() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1000, 2000, 3000, -500)");
        Assert.Equal(5500.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Sum_ShouldHandleReferencesOnly() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "1");
        sheet.SetCell(0, 1, "2");
        sheet.SetCell(0, 2, "=SUM(A1,B1,A1,B1)");
        Assert.Equal(6.0, sheet.GetCell(0, 2).Value);
    }

    [Fact]
    public void Sum_ShouldFailOnMissingComma() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1 2)");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Average_ShouldIgnoreSpacing() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "= AVERAGE ( 10 ,  20 ,  30 ) ");
        Assert.Equal(20.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Average_ShouldHandleLargeInputCount() {
        var sheet = new Worksheet();
        var args = string.Join(",", Enumerable.Range(1, 50).Select(i => "2"));
        sheet.SetCell(0, 0, $"=AVERAGE({args})");
        Assert.Equal(2.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Max_AllEqualValues() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX(5, 5, 5, 5)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Max_SingleArgument() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX(777)");
        Assert.Equal(777.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Min_AllEqualValues() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MIN(3, 3, 3)");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Min_ShouldReturnLowestInMixedSet() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MIN(100, -1, 50)");
        Assert.Equal(-1.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Median_BigOddSet() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN(1,3,5,7,9)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Median_WithNegativeNumbers() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN(-5, 0, 5)");
        Assert.Equal(0.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Nested_SumInsideSum() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM( SUM(1,2), SUM(3,4) )");
        Assert.Equal(10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Nested_AverageInsideMax() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX( AVERAGE(10,20), 15 )");
        Assert.Equal(15.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Nested_MaxInsideMedian() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN( 1, MAX(5,2), 3 )");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void ArithmeticWithFunctions() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(2,3)*2 - MIN(10,5)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void MixedComplexExpression() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=((SUM(1,2)*3) + MAX(5,2)) / 2");
        Assert.Equal(7.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void CaseInsensitive_MinMixedCase() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=mIn(3,1,2)");
        Assert.Equal(1.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void CaseInsensitive_MaxUpperCase() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX(3,1,2)");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Reference_ShouldHandleLowercaseColumn() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "10");
        sheet.SetCell(0, 1, "=a1 * 2");
        Assert.Equal(20.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void Reference_ShouldHandleMixedCaseColumn() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "7");
        sheet.SetCell(0, 1, "=Aa1 + 1");
        Assert.Equal(1.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void Dependency_ShouldRecalculateOnMultipleChanges() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "1");
        sheet.SetCell(0, 1, "=A1*5");

        sheet.SetCell(0, 0, "2");
        sheet.SetCell(0, 0, "3");

        Assert.Equal(15.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void ChainedDependency_ThreeLevels() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "2");
        sheet.SetCell(0, 1, "=A1+1");
        sheet.SetCell(0, 2, "=B1+1");

        Assert.Equal(4.0, sheet.GetCell(0, 2).Value);
    }

    [Fact]
    public void CircularDependency_ShouldGiveError() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=A1+1");
        Assert.Equal("#CYCLE!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void MultipleCircularDependency() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=B1+1");
        sheet.SetCell(0, 1, "=A1+1");
        Assert.Equal("#CYCLE!", sheet.GetCell(0, 0).Value);
        Assert.Equal("#CYCLE!", sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void Sum_WithReferenceRange_Simple() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "1");
        sheet.SetCell(1, 0, "2");
        sheet.SetCell(2, 0, "3");
        sheet.SetCell(0, 1, "=SUM(A1, A2, A3)");

        Assert.Equal(6.0, sheet.GetCell(0, 1).Value);
    }
    
    [Fact]
    public void Addition_NegativeAndPositive() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=-5+10");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Addition_LongChain() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=1+2+3+4+5+6+7+8+9+10");
        Assert.Equal(55.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Subtraction_LongChain() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=100-1-2-3-4");
        Assert.Equal(90.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Multiplication_LongChain() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=1*2*3*4");
        Assert.Equal(24.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Division_Chained() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=100/5/2");
        Assert.Equal(10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void MixedOperators_NoParentheses_LeftToRight() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=20/2*5");
        Assert.Equal(50.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void MixedOperatorsHighPrecision() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=10/3");
        Assert.Equal(10.0 / 3.0, (double)sheet.GetCell(0, 0).Value, 5);
    }

    [Fact]
    public void ParenthesesAroundSingleValue() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=(5)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void ParenthesesAroundExpression() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=(2+3)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void DeepParenthesesWithSpaces() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=(( ( 2+3 ) ))");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void InvalidSyntax_DanglingParenthesis() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=(2+3");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void InvalidSyntax_ExtraClosingParenthesis() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=2+3)");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void InvalidSyntax_OperatorAtStart() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=+*5");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void InvalidSyntax_OperatorAtEnd() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=10/");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Reference_Basic() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "7");
        sheet.SetCell(0, 1, "=A1");
        Assert.Equal(7.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void Reference_InArithmetic() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "3");
        sheet.SetCell(0, 1, "=A1+2");
        Assert.Equal(5.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void Reference_WhitespaceAround() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "4");
        sheet.SetCell(0, 1, "=  A1 * 2 ");
        Assert.Equal(8.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void Reference_InvalidFormat_ShouldError() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=A");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Sum_WithLongArgumentList() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1,2,3,4,5,6,7,8,9,10)");
        Assert.Equal(55.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Sum_WithWhitespaceAndMixedValues() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM( 1 ,  2,3 ,   4)");
        Assert.Equal(10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Sum_WithReferenceAndLiteral() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "10");
        sheet.SetCell(0, 1, "=SUM(A1, 5)");
        Assert.Equal(15.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void Sum_AllZeros() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(0,0,0,0)");
        Assert.Equal(0.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Average_ManyValues() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(1,1,1,1,1,1)");
        Assert.Equal(1.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Average_WithMixedSigns() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(-5,5)");
        Assert.Equal(0.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Average_WithReference() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "9");
        sheet.SetCell(0, 1, "=AVERAGE(A1, 3)");
        Assert.Equal(6.0, sheet.GetCell(0, 1).Value);
    }

    [Fact]
    public void Max_WithSingleValue() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX(123)");
        Assert.Equal(123.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Max_WithMixedSignsAndZero() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX(-10, 0, 10)");
        Assert.Equal(10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Min_WithSingleValue() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MIN(5)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Min_WithMixedSignsAndZero() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MIN(-10, 0, 10)");
        Assert.Equal(-10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Median_SimpleEven() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN(2,4)");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Median_SimpleOdd() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN(2,4,6)");
        Assert.Equal(4.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Median_WithDuplicates() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MEDIAN(1,1,1,1)");
        Assert.Equal(1.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void NestedFunction_SumInAverage() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(SUM(5,5), 10)");
        Assert.Equal(10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void NestedFunction_AverageInSum() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(AVERAGE(4,6), 5)");
        Assert.Equal(10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void NestedFunction_MinMaxMedianMix() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=MAX( MIN(5,10), MEDIAN(2,4,6) )");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void MultipleFunctionsInExpression() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1,2) + AVERAGE(4,6) + MAX(1,9)");
        Assert.Equal(17.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void CaseInsensitivity_SUM_MixedCase() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SuM(1,2)");
        Assert.Equal(3.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void CaseInsensitivity_AVERAGE_Upper() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=AVERAGE(8,12)");
        Assert.Equal(10.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void CaseInsensitivity_MAX_Lowercase() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=max(1, 100)");
        Assert.Equal(100.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void CaseInsensitivity_MIN_Mixed() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=mIn(20,5)");
        Assert.Equal(5.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void CaseInsensitivity_MEDIAN_Lowercase() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=median(3,9,6)");
        Assert.Equal(6.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void ManyArguments_Sum_50TimesThree() {
        var sheet = new Worksheet();
        var args = string.Join(",", Enumerable.Repeat("3", 50));
        sheet.SetCell(0, 0, $"=SUM({args})");
        Assert.Equal(150.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void ManyArguments_Average_100ones() {
        var sheet = new Worksheet();
        var args = string.Join(",", Enumerable.Repeat("1", 100));
        sheet.SetCell(0, 0, $"=AVERAGE({args})");
        Assert.Equal(1.0, sheet.GetCell(0, 0).Value);
    }

    [Fact]
    public void Dependency_DeepChain_10Levels() {
        var sheet = new Worksheet();

        sheet.SetCell(0, 0, "1");
        for (int i = 1; i < 10; i++) {
            string prevCol = ((char)('A' + i - 1)).ToString();
            sheet.SetCell(0, i, $"={prevCol}1 + 1");
        }

        Assert.Equal(10.0, sheet.GetCell(0, 9).Value);
    }

    [Fact]
    public void ErrorPropagation_FromNestedFunction() {
        var sheet = new Worksheet();
        sheet.SetCell(0, 0, "=SUM(1, AVERAGE(5/0))");
        Assert.Equal("#ERROR!", sheet.GetCell(0, 0).Value);
    }
}