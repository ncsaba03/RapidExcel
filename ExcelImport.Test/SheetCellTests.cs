using ExcelImport.Spreadsheet;

namespace ExcelImport.Test;

public class SheetCellTests
{
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
    public void SheetCell_AddColumns(string start, int col, string expected)
    {
        SheetCell cell = SheetCell.Parse(start).AddColumns(col);
        Assert.Equal(expected, cell.ToString());
    }
}
