using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Test;

/// <summary>
/// Helper class to create test Excel files and OpenXML structures
/// </summary>
public static class TestDataBuilder
{
    /// <summary>
    /// Creates a simple Excel file in memory with the specified data
    /// </summary>
    public static string CreateSimpleExcelFile(string[][] data, bool includeSharedStrings = true)
    {
        var filePath = Path.GetTempFileName() + ".xlsx";

        using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            // Add workbook part
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add worksheet part
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add sheets collection
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            };
            sheets.Append(sheet);

            // Add shared string table if needed
            SharedStringTablePart? sharedStringPart = null;
            SharedStringTable? sharedStringTable = null;

            if (includeSharedStrings)
            {
                sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
                sharedStringTable = new SharedStringTable();
                sharedStringPart.SharedStringTable = sharedStringTable;
            }

            // Add data to worksheet
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            uint rowIndex = 1;

            foreach (var rowData in data)
            {
                var row = new Row { RowIndex = rowIndex };
                uint colIndex = 1;

                foreach (var cellValue in rowData)
                {
                    var cellRef = GetCellReference(colIndex, rowIndex);
                    var cell = new Cell { CellReference = cellRef };

                    if (includeSharedStrings && !string.IsNullOrEmpty(cellValue))
                    {
                        // Add to shared strings
                        var index = AddSharedString(sharedStringTable!, cellValue);
                        cell.CellValue = new CellValue(index.ToString());
                        cell.DataType = CellValues.SharedString;
                    }
                    else
                    {
                        cell.CellValue = new CellValue(cellValue);
                        cell.DataType = CellValues.String;
                    }

                    row.Append(cell);
                    colIndex++;
                }

                sheetData.Append(row);
                rowIndex++;
            }

            // Add sheet dimension
            var dimension = $"A1:{GetCellReference((uint)data[0].Length, (uint)data.Length)}";
            worksheetPart.Worksheet.InsertBefore(
                new SheetDimension { Reference = dimension },
                sheetData);

            workbookPart.Workbook.Save();
        }

        return filePath;
    }

    /// <summary>
    /// Creates an Excel file with only numeric values (no shared strings needed)
    /// </summary>
    public static string CreateNumericExcelFile()
    {
        var data = new string[][]
        {
            ["Number1", "Number2"],
            ["123", "456"],
            ["789", "012"]
        };

        return CreateSimpleExcelFile(data, includeSharedStrings: false);
    }

    /// <summary>
    /// Creates an Excel file without any worksheets (invalid)
    /// </summary>
    public static string CreateInvalidExcelFile()
    {
        var filePath = Path.GetTempFileName() + ".xlsx";

        using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            workbookPart.Workbook.Save();
        }

        return filePath;
    }

    /// <summary>
    /// Creates an Excel file for testing SimpleTestModel import
    /// </summary>
    public static string CreateSimpleTestModelFile()
    {
        var data = new string[][]
        {
            ["Name", "Age", "Score"],
            ["John", "30", "95.5"],
            ["Jane", "25", "87.3"],
            ["Bob", "35", "91.2"]
        };

        return CreateSimpleExcelFile(data);
    }

    /// <summary>
    /// Creates an Excel file with different header row index
    /// </summary>
    public static string CreateFileWithHeaderAtRow(uint headerRow)
    {
        var rows = new List<string[]>();

        // Add empty rows before header
        for (uint i = 0; i < headerRow; i++)
        {
            rows.Add(["", "", ""]);
        }

        // Add header and data
        rows.AddRange([
            ["Name", "Age", "Score"],
            ["John", "30", "95.5"]
        ]);

        return CreateSimpleExcelFile(rows.ToArray());
    }

    /// <summary>
    /// Creates an Excel file with missing required fields
    /// </summary>
    public static string CreateFileWithMissingData()
    {
        var data = new string[][]
        {
            ["Required", "Optional"],
            ["HasValue", "AlsoHasValue"],
            ["", "HasValue"] // Missing required field
        };

        return CreateSimpleExcelFile(data);
    }

    private static int AddSharedString(SharedStringTable sharedStringTable, string text)
    {
        var index = 0;
        foreach (var item in sharedStringTable.Elements<SharedStringItem>())
        {
            if (item.InnerText == text)
                return index;
            index++;
        }

        sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
        return index;
    }

    private static string GetCellReference(uint colIndex, uint rowIndex)
    {
        var columnLetter = GetColumnLetter(colIndex);
        return $"{columnLetter}{rowIndex}";
    }

    private static string GetColumnLetter(uint colIndex)
    {
        string columnLetter = "";
        while (colIndex > 0)
        {
            var modulo = (colIndex - 1) % 26;
            columnLetter = Convert.ToChar(65 + modulo) + columnLetter;
            colIndex = (colIndex - modulo) / 26;
        }
        return columnLetter;
    }
}
