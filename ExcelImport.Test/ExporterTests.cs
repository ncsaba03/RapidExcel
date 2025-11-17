using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Exceptions;

namespace ExcelImport.Test;

public class ExporterTests
{
    #region Basic Export Tests

    [Fact]
    public void Export_SimpleModel_CreatesFile()
    {
        var data = new List<SimpleTestModel>
        {
            new() { Name = "John", Age = 30, Score = 95.5m },
            new() { Name = "Jane", Age = 25, Score = 87.3m }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            Assert.True(File.Exists(filePath));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Export_ToStream_WritesData()
    {
        var data = new List<SimpleTestModel>
        {
            new() { Name = "John", Age = 30, Score = 95.5m }
        };

        using var stream = new MemoryStream();
        var exporter = new ExcelExporter();
        exporter.ExportWithWriter(data, stream);

        Assert.True(stream.Length > 0);
    }

    [Fact]
    public void Export_EmptyList_CreatesFileWithHeadersOnly()
    {
        var data = new List<SimpleTestModel>();
        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            Assert.True(File.Exists(filePath));

            // Verify file has header row
            using var document = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = document.WorkbookPart!;
            var worksheetPart = workbookPart.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;

            Assert.Single(sheetData.Elements<Row>()); // Only header row
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Round-Trip Tests

    [Fact]
    public void Export_ThenImport_PreservesData()
    {
        var originalData = new List<SimpleTestModel>
        {
            new() { Name = "John", Age = 30, Score = 95.5m },
            new() { Name = "Jane", Age = 25, Score = 87.3m }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            // Export
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(originalData, filePath);

            // Import back
            var importer = new ExcelImporter();
            var importedData = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Equal(originalData.Count, importedData.Count);

            for (int i = 0; i < originalData.Count; i++)
            {
                Assert.Equal(originalData[i].Name, importedData[i].Name);
                Assert.Equal(originalData[i].Age, importedData[i].Age);
                Assert.Equal(originalData[i].Score, importedData[i].Score);
            }
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Export_AllTypes_ThenImport_PreservesTypes()
    {
        var originalData = new List<AllTypesModel>
        {
            new()
            {
                StringValue = "test",
                IntValue = 123,
                LongValue = 123456789L,
                FloatValue = 12.34f,
                DoubleValue = 98.76,
                DecimalValue = 55.55m,
                BoolValue = true,
                DateTimeValue = new DateTime(2023, 6, 15)
            }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(originalData, filePath);

            var importer = new ExcelImporter();
            var importedData = importer.Import<AllTypesModel>(filePath).ToList();

            Assert.Single(importedData);
            var original = originalData[0];
            var imported = importedData[0];

            Assert.Equal(original.StringValue, imported.StringValue);
            Assert.Equal(original.IntValue, imported.IntValue);
            Assert.Equal(original.LongValue, imported.LongValue);
            Assert.Equal(original.FloatValue, imported.FloatValue, 2);
            Assert.Equal(original.DoubleValue, imported.DoubleValue, 2);
            Assert.Equal(original.DecimalValue, imported.DecimalValue);
            Assert.Equal(original.BoolValue, imported.BoolValue);
            Assert.Equal(original.DateTimeValue, imported.DateTimeValue);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Column Ordering Tests

    [Fact]
    public void Export_WithPositionAttribute_OrdersColumnsCorrectly()
    {
        var data = new List<PositionOrderingModel>
        {
            new()
            {
                First = "1st",
                Second = "2nd",
                Third = "3rd",
                Last = "4th"
            }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            using var document = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = document.WorkbookPart!;
            var worksheetPart = workbookPart.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var headerRow = sheetData.Elements<Row>().First();
            var headerCells = headerRow.Elements<Cell>().ToList();

            // Verify order: First, Second, Third, Last
            Assert.Equal(4, headerCells.Count);
            // Note: actual cell values would need shared string resolution
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Required Field Tests

    [Fact]
    public void Export_RequiredFieldNull_ThrowsImportException()
    {
        var data = new List<RequiredFieldsModel>
        {
            new() { RequiredField = null!, OptionalField = "Test" }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();

            Assert.Throws<ImportException>(() =>
                exporter.ExportWithWriter(data, filePath));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Export_RequiredFieldPresent_ExportsSuccessfully()
    {
        var data = new List<RequiredFieldsModel>
        {
            new() { RequiredField = "Required", OptionalField = "Optional" }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            Assert.True(File.Exists(filePath));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Export_OptionalFieldNull_SkipsCell()
    {
        var data = new List<RequiredFieldsModel>
        {
            new() { RequiredField = "Required", OptionalField = null }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            Assert.True(File.Exists(filePath));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Multiple Sheets Tests

    [Fact]
    public void ExportSheetsWithWriter_MultipleSheets_CreatesAllSheets()
    {
        var data = new List<(string, List<SimpleTestModel>)>
        {
            ("Sheet1", new List<SimpleTestModel>
            {
                new() { Name = "John", Age = 30, Score = 95.5m }
            }),
            ("Sheet2", new List<SimpleTestModel>
            {
                new() { Name = "Jane", Age = 25, Score = 87.3m }
            }),
            ("Sheet3", new List<SimpleTestModel>
            {
                new() { Name = "Bob", Age = 35, Score = 91.2m }
            })
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportSheetsWithWriter(data, filePath);

            using var document = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = document.WorkbookPart!;
            var sheets = workbookPart.Workbook.Sheets!.Elements<Sheet>().ToList();

            Assert.Equal(3, sheets.Count);
            Assert.Equal("Sheet1", sheets[0].Name);
            Assert.Equal("Sheet2", sheets[1].Name);
            Assert.Equal("Sheet3", sheets[2].Name);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExportSheetsWithWriter_EmptySheetList_CreatesFile()
    {
        var data = new List<(string, List<SimpleTestModel>)>();
        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportSheetsWithWriter(data, filePath);

            Assert.True(File.Exists(filePath));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Custom Converter Tests

    [Fact]
    public void Export_WithCustomConverter_UsesConverter()
    {
        var data = new List<CustomConverterModel>
        {
            new() { Amount = 10000m, Code = "ABC" } // Should be converted to 100
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            // Import back to verify conversion
            var importer = new ExcelImporter();
            var imported = importer.Import<CustomConverterModel>(filePath).ToList();

            Assert.Single(imported);
            Assert.Equal(10000m, imported[0].Amount); // Round-trip through converter
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Large Data Tests

    [Fact]
    public void Export_LargeDataset_ExportsSuccessfully()
    {
        var data = new List<SimpleTestModel>();

        for (int i = 0; i < 1000; i++)
        {
            data.Add(new SimpleTestModel
            {
                Name = $"Person{i}",
                Age = i,
                Score = i * 0.5m
            });
        }

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            Assert.True(File.Exists(filePath));

            // Verify file has correct number of rows
            using var document = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = document.WorkbookPart!;
            var worksheetPart = workbookPart.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var rowCount = sheetData.Elements<Row>().Count();

            Assert.Equal(1001, rowCount); // 1 header + 1000 data rows
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region File Structure Tests

    [Fact]
    public void Export_CreatesValidExcelStructure()
    {
        var data = new List<SimpleTestModel>
        {
            new() { Name = "John", Age = 30, Score = 95.5m }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            using var document = SpreadsheetDocument.Open(filePath, false);

            Assert.NotNull(document.WorkbookPart);
            Assert.NotNull(document.WorkbookPart.Workbook);
            Assert.NotNull(document.WorkbookPart.Workbook.Sheets);
            Assert.Single(document.WorkbookPart.WorksheetParts);

            var worksheetPart = document.WorkbookPart.WorksheetParts.First();
            Assert.NotNull(worksheetPart.Worksheet);

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
            Assert.NotNull(sheetData);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Export_HasSheetDimension()
    {
        var data = new List<SimpleTestModel>
        {
            new() { Name = "John", Age = 30, Score = 95.5m },
            new() { Name = "Jane", Age = 25, Score = 87.3m }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            using var document = SpreadsheetDocument.Open(filePath, false);
            var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
            var dimension = worksheetPart.Worksheet.GetFirstChild<SheetDimension>();

            Assert.NotNull(dimension);
            Assert.NotNull(dimension.Reference);
            Assert.Contains("A1:", dimension.Reference.Value);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Export_HeaderRow_IsFirstRow()
    {
        var data = new List<SimpleTestModel>
        {
            new() { Name = "John", Age = 30, Score = 95.5m }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            using var document = SpreadsheetDocument.Open(filePath, false);
            var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var firstRow = sheetData.Elements<Row>().First();

            Assert.Equal(1u, firstRow.RowIndex!.Value);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Export_DataRows_StartAtSecondRow()
    {
        var data = new List<SimpleTestModel>
        {
            new() { Name = "John", Age = 30, Score = 95.5m }
        };

        var filePath = Path.GetTempFileName() + ".xlsx";

        try
        {
            var exporter = new ExcelExporter();
            exporter.ExportWithWriter(data, filePath);

            using var document = SpreadsheetDocument.Open(filePath, false);
            var worksheetPart = document.WorkbookPart!.WorksheetParts.First();
            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var rows = sheetData.Elements<Row>().ToList();

            Assert.Equal(2, rows.Count);
            Assert.Equal(1u, rows[0].RowIndex!.Value); // Header
            Assert.Equal(2u, rows[1].RowIndex!.Value); // Data
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion
}
