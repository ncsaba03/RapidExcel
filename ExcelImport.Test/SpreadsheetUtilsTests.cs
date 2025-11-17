using ExcelImport.Spreadsheet;

namespace ExcelImport.Test;

public class SpreadsheetUtilsTests
{
    #region SheetCell Tests

    [Theory]
    [InlineData("A1", "A", 1)]
    [InlineData("Z10", "Z", 10)]
    [InlineData("AA1", "AA", 1)]
    [InlineData("ZZ99", "ZZ", 99)]
    [InlineData("AAA100", "AAA", 100)]
    public void SheetCell_Parse_ValidInput_ParsesCorrectly(string input, string expectedCol, int expectedRow)
    {
        var cell = SheetCell.Parse(input);

        Assert.Equal(expectedCol, cell.Col);
        Assert.Equal(expectedRow, cell.Row);
    }

    [Theory]
    [InlineData("A", "A", 1)]
    public void SheetCell_Constructor_SetsProperties(string col, string expectedCol, int row)
    {
        var cell = new SheetCell(col, row);

        Assert.Equal(expectedCol, cell.Col);
        Assert.Equal(row, cell.Row);
    }

    [Theory]
    [InlineData("A10", 0, "A10")]
    [InlineData("A10", 10, "K10")]
    [InlineData("A10", 30, "AE10")]
    [InlineData("D10", 22, "Z10")]
    [InlineData("D10", 23, "AA10")]
    [InlineData("AA10", 10, "AK10")]
    [InlineData("D10", -3, "A10")]
    [InlineData("AA10", 26, "BA10")]
    [InlineData("BA10", 649, "ZZ10")]
    [InlineData("ZZ10", 1, "AAA10")]
    [InlineData("ZZ10", 1379, "CBA10")]
    [InlineData("A10", 1, "B10")]
    public void SheetCell_AddColumns_ReturnsCorrectCell(string start, int col, string expected)
    {
        var cell = SheetCell.Parse(start).AddColumns(col);

        Assert.Equal(expected, cell.ToString());
    }

    [Theory]
    [InlineData("A1", 0, "A1")]
    [InlineData("A1", 1, "A2")]
    [InlineData("A1", 10, "A11")]
    [InlineData("A10", -5, "A5")]
    [InlineData("B5", 100, "B105")]
    public void SheetCell_AddRows_ReturnsCorrectCell(string start, int rows, string expected)
    {
        var cell = SheetCell.Parse(start).AddRows(rows);

        Assert.Equal(expected, cell.ToString());
    }

    [Theory]
    [InlineData("A", 1)]
    [InlineData("B", 2)]
    [InlineData("Z", 26)]
    [InlineData("AA", 27)]
    [InlineData("AB", 28)]
    [InlineData("AZ", 52)]
    [InlineData("BA", 53)]
    [InlineData("ZZ", 702)]
    [InlineData("AAA", 703)]
    public void SheetCell_ColIndex_ReturnsCorrectIndex(string col, int expectedIndex)
    {
        var cell = new SheetCell(col, 1);

        Assert.Equal(expectedIndex, cell.ColIndex);
    }

    [Fact]
    public void SheetCell_Equals_SameCells_ReturnsTrue()
    {
        var cell1 = new SheetCell("A", 1);
        var cell2 = new SheetCell("A", 1);

        Assert.True(cell1.Equals(cell2));
    }

    [Fact]
    public void SheetCell_Equals_DifferentCells_ReturnsFalse()
    {
        var cell1 = new SheetCell("A", 1);
        var cell2 = new SheetCell("B", 1);

        Assert.False(cell1.Equals(cell2));
    }

    [Fact]
    public void SheetCell_GetHashCode_SameCells_ReturnsSameHash()
    {
        var cell1 = new SheetCell("A", 1);
        var cell2 = new SheetCell("A", 1);

        Assert.Equal(cell1.GetHashCode(), cell2.GetHashCode());
    }

    [Theory]
    [InlineData("A1", "A2", -1)]
    [InlineData("A2", "A1", 1)]
    [InlineData("A1", "B1", -1)]
    [InlineData("B1", "A1", 1)]
    [InlineData("A1", "A1", 0)]
    public void SheetCell_CompareTo_ReturnsCorrectOrder(string cell1Str, string cell2Str, int expectedSign)
    {
        var cell1 = SheetCell.Parse(cell1Str);
        var cell2 = SheetCell.Parse(cell2Str);

        var result = cell1.CompareTo(cell2);

        if (expectedSign < 0)
            Assert.True(result < 0);
        else if (expectedSign > 0)
            Assert.True(result > 0);
        else
            Assert.Equal(0, result);
    }

    [Fact]
    public void SheetCell_ToString_ReturnsCorrectFormat()
    {
        var cell = new SheetCell("AB", 123);

        Assert.Equal("AB123", cell.ToString());
    }

    #endregion

    #region SheetHelper Tests

    [Theory]
    [InlineData("A", 1)]
    [InlineData("B", 2)]
    [InlineData("Z", 26)]
    [InlineData("AA", 27)]
    [InlineData("AB", 28)]
    [InlineData("AZ", 52)]
    [InlineData("BA", 53)]
    [InlineData("ZZ", 702)]
    [InlineData("AAA", 703)]
    [InlineData("AAB", 704)]
    public void SheetHelper_GetColumnIndex_ValidColumn_ReturnsCorrectIndex(string col, int expected)
    {
        var result = SheetHelper.GetColumnIndex(col.AsSpan());

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(1u, "A")]
    [InlineData(2u, "B")]
    [InlineData(26u, "Z")]
    [InlineData(27u, "AA")]
    [InlineData(28u, "AB")]
    [InlineData(52u, "AZ")]
    [InlineData(53u, "BA")]
    [InlineData(702u, "ZZ")]
    [InlineData(703u, "AAA")]
    [InlineData(704u, "AAB")]
    public void SheetHelper_TransformToCharacterIndex_ValidIndex_ReturnsCorrectColumn(uint index, string expected)
    {
        var result = SheetHelper.TransformToCharacterIndex(index);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("A1", "A")]
    [InlineData("Z10", "Z")]
    [InlineData("AA1", "AA")]
    [InlineData("ZZ99", "ZZ")]
    [InlineData("AAA100", "AAA")]
    public void SheetHelper_GetColumnIndexFromCellReference_ValidReference_ReturnsCorrectIndex(string cellRef, string expected)
    {
        var result = SheetHelper.GetColumnIndexFromCellReference(cellRef.AsSpan()).ToString();

        Assert.Equal(expected, result);
    }

    [Fact]
    public void SheetHelper_RoundTrip_IndexToColumnToIndex_PreservesValue()
    {
        for (uint i = 1; i <= 1000; i++)
        {
            var col = SheetHelper.TransformToCharacterIndex(i);
            var index = SheetHelper.GetColumnIndex(col.AsSpan());

            Assert.Equal(i, (uint)index);
        }
    }

    #endregion

    #region SheetRange Tests

    [Theory]
    [InlineData("A1:B10", "A1", "B10")]
    [InlineData("AA1:ZZ100", "AA1", "ZZ100")]
    [InlineData("C5:E20", "C5", "E20")]
    public void SheetRange_Parse_ValidRange_ParsesCorrectly(string input, string expectedFrom, string expectedTo)
    {
        var range = SheetRange.Parse(input.AsSpan());

        Assert.Equal(expectedFrom, range.From.ToString());
        Assert.Equal(expectedTo, range.To.ToString());
    }

    [Theory]
    [InlineData("A1:B10")]
    [InlineData("AA1:ZZ100")]
    [InlineData("C5:E20")]
    public void SheetRange_IsValidForRange_ValidRange_ReturnsTrue(string input)
    {
        var result = SheetRange.IsValidForRange(input.AsSpan());

        Assert.True(result);
    }

    [Theory]
    [InlineData("A1B10")]
    [InlineData("A1:")]
    [InlineData(":B10")]
    [InlineData("invalid")]
    [InlineData("")]
    public void SheetRange_IsValidForRange_InvalidRange_ReturnsFalse(string input)
    {
        var result = SheetRange.IsValidForRange(input.AsSpan());

        Assert.False(result);
    }

    [Fact]
    public void SheetRange_RowLength_CalculatesCorrectly()
    {
        var range = SheetRange.Parse("A1:A10");

        Assert.Equal(10, range.RowLength);
    }

    [Fact]
    public void SheetRange_ColumnLength_CalculatesCorrectly()
    {
        var range = SheetRange.Parse("A1:Z1");

        Assert.Equal(26, range.CoulumnLength); // Note: typo in property name
    }

    [Theory]
    [InlineData("A1:B10", 1, "A2:B11")]
    [InlineData("C5:E20", 5, "C10:E25")]
    public void SheetRange_AddRowsToBegining_AddsRowsToStart(string input, int rows, string expected)
    {
        var range = SheetRange.Parse(input.AsSpan());

        range = range.AddRowsToBegining(rows);

        Assert.Equal(expected, range.ToString());
    }

    [Theory]
    [InlineData("A1:B10", 1, "A1:B11")]
    [InlineData("C5:E20", 5, "C5:E25")]
    public void SheetRange_AddRowsToEnd_AddsRowsToEnd(string input, int rows, string expected)
    {
        var range = SheetRange.Parse(input.AsSpan());

        range = range.AddRowsToEnd(rows);

        Assert.Equal(expected, range.ToString());
    }

    [Theory]
    [InlineData("A1:B10", 1, "B1:C10")]
    [InlineData("C5:E20", 2, "E5:G20")]
    public void SheetRange_AddColumnsToBegining_AddsColumnsToStart(string input, int cols, string expected)
    {
        var range = SheetRange.Parse(input.AsSpan());

        range = range.AddColumnsToBegining(cols);

        Assert.Equal(expected, range.ToString());
    }

    [Theory]
    [InlineData("A1:B10", 1, "A1:C10")]
    [InlineData("C5:E20", 2, "C5:G20")]
    public void SheetRange_AddColumnsToEnd_AddsColumnsToEnd(string input, int cols, string expected)
    {
        var range = SheetRange.Parse(input.AsSpan());

        range = range.AddColumnsToEnd(cols);

        Assert.Equal(expected, range.ToString());
    }

    [Fact]
    public void SheetRange_SetFrom_UpdatesFromCell()
    {
        var range = SheetRange.Parse("A1:B10");

        range = range.SetFrom("C5");

        Assert.Equal("C5", range.From.ToString());
    }

    [Fact]
    public void SheetRange_SetTo_UpdatesToCell()
    {
        var range = SheetRange.Parse("A1:B10");

        range = range.SetTo("Z99");

        Assert.Equal("Z99", range.To.ToString());
    }

    [Theory]
    [InlineData("A1:B10", "A1", true)]
    [InlineData("A1:B10", "B10", true)]
    [InlineData("A1:B10", "A5", true)]
    [InlineData("A1:B10", "C1", false)]
    [InlineData("A1:B10", "A11", false)]
    public void SheetRange_IsInRange_ChecksCellCorrectly(string rangeStr, string cellStr, bool expected)
    {
        var range = SheetRange.Parse(rangeStr.AsSpan());
        var cell = SheetCell.Parse(cellStr);

        var result = range.IsInRange(cell);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("A1:B10", "A1", true)]
    [InlineData("A1:B10", "B10", true)]
    [InlineData("A1:B10", "C1", false)]
    public void SheetRange_IsInRange_WithString_ChecksCellCorrectly(string rangeStr, string cellStr, bool expected)
    {
        var range = SheetRange.Parse(rangeStr.AsSpan());

        var result = range.IsInRange(cellStr.AsSpan());

        Assert.Equal(expected, result);
    }

    [Fact]
    public void SheetRange_Equals_SameRanges_ReturnsTrue()
    {
        var range1 = SheetRange.Parse("A1:B10");
        var range2 = SheetRange.Parse("A1:B10");

        Assert.True(range1.Equals(range2));
    }

    [Fact]
    public void SheetRange_Equals_DifferentRanges_ReturnsFalse()
    {
        var range1 = SheetRange.Parse("A1:B10");
        var range2 = SheetRange.Parse("C1:D10");

        Assert.False(range1.Equals(range2));
    }

    [Fact]
    public void SheetRange_ToString_ReturnsCorrectFormat()
    {
        var range = SheetRange.Parse("A1:B10");

        Assert.Equal("A1:B10", range.ToString());
    }

    [Fact]
    public void SheetRange_CompareTo_ReturnsCorrectOrder()
    {
        var range1 = SheetRange.Parse("A1:B10");
        var range2 = SheetRange.Parse("C1:D10");

        var result = range1.CompareTo(range2);

        Assert.True(result != 0); // Different ranges should not be equal
    }

    #endregion
}
