using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using RapidExcel.Utils;

namespace RapidExcel.Test;

public class FastSharedStringReaderTests
{
    #region Basic Functionality Tests

    [Fact]
    public void Create_WithValidSharedStringTable_CreatesReader()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["Apple", "Banana", "Cherry"]);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;

            // Act
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Assert
            Assert.NotNull(reader);
            Assert.Equal(3, reader.Count);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_ValidIndex_ReturnsCorrectString()
    {
        // Arrange
        var strings = new[] { "First", "Second", "Third", "Fourth" };
        var filePath = CreateTestFileWithSharedStrings(strings);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act & Assert
            for (int i = 0; i < strings.Length; i++)
            {
                var result = reader.GetString(i);
                Assert.Equal(strings[i], result);
            }
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_MultipleCallsSameIndex_ReturnsSameString()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["Test"]);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act
            var result1 = reader.GetString(0);
            var result2 = reader.GetString(0);
            var result3 = reader.GetString(0);

            // Assert
            Assert.Equal("Test", result1);
            Assert.Equal("Test", result2);
            Assert.Equal("Test", result3);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Boundary Tests

    [Fact]
    public void GetString_NegativeIndex_ThrowsArgumentOutOfRangeException()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["Test"]);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act & Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetString(-1));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_IndexEqualToCount_ThrowsArgumentOutOfRangeException()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["First", "Second"]);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act & Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetString(2));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_IndexGreaterThanCount_ThrowsArgumentOutOfRangeException()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["Test"]);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act & Assert
            Assert.Throws<ArgumentOutOfRangeException>(() => reader.GetString(100));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Special Characters and Encoding Tests

    [Fact]
    public void GetString_WithUnicodeCharacters_ReturnsCorrectString()
    {
        // Arrange
        var strings = new[] { "Hello 世界", "Привет мир", "مرحبا العالم", "Árvíztűrő tükörfúrógép" };
        var filePath = CreateTestFileWithSharedStrings(strings);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act & Assert
            for (int i = 0; i < strings.Length; i++)
            {
                var result = reader.GetString(i);
                Assert.Equal(strings[i], result);
            }
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_WithSpecialCharacters_ReturnsCorrectString()
    {
        // Arrange
        var strings = new[]
        {
            "Line1\nLine2",
            "Tab\tSeparated",
            "Quote\"Test",
            "<xml>&amp;</xml>",
            "Special: !@#$%^&*()"
        };
        var filePath = CreateTestFileWithSharedStrings(strings);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act & Assert
            for (int i = 0; i < strings.Length; i++)
            {
                var result = reader.GetString(i);
                Assert.Equal(strings[i], result);
            }
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_WithEmptyString_ReturnsEmptyString()
    {
        // Arrange
        var strings = new[] { "Before", "", "After" };
        var filePath = CreateTestFileWithSharedStrings(strings);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act
            var result = reader.GetString(1);

            // Assert
            Assert.Equal(string.Empty, result);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_WithVeryLongString_ReturnsCorrectString()
    {
        // Arrange
        var longString = new string('A', 10000);
        var strings = new[] { longString };
        var filePath = CreateTestFileWithSharedStrings(strings);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act
            var result = reader.GetString(0);

            // Assert
            Assert.Equal(longString, result);
            Assert.Equal(10000, result.Length);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Large Dataset Tests

    [Fact]
    public void Create_WithLargeNumberOfStrings_HandlesCorrectly()
    {
        // Arrange
        var count = 10000;
        var strings = Enumerable.Range(0, count)
            .Select(i => $"String_{i}")
            .ToArray();
        var filePath = CreateTestFileWithSharedStrings(strings);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Assert
            Assert.Equal(count, reader.Count);

            // Verify random access
            Assert.Equal("String_0", reader.GetString(0));
            Assert.Equal("String_5000", reader.GetString(5000));
            Assert.Equal("String_9999", reader.GetString(9999));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_RandomAccessPattern_PerformsWell()
    {
        // Arrange
        var count = 1000;
        var strings = Enumerable.Range(0, count)
            .Select(i => $"Item_{i}")
            .ToArray();
        var filePath = CreateTestFileWithSharedStrings(strings);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act - Access in random order
            var random = new Random(42); // Fixed seed for reproducibility
            for (int i = 0; i < 1000; i++)
            {
                int index = random.Next(count);
                var result = reader.GetString(index);

                // Assert
                Assert.Equal($"Item_{index}", result);
            }
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Dispose Tests

    [Fact]
    public void Dispose_CalledOnce_CleansUpResources()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["Test"]);
        string? tempFilePath = null;

        try
        {
            using (var document = SpreadsheetDocument.Open(filePath, false))
            {
                var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
                var reader = FastSharedStringReader.Create(sharedStringPart);

                
                var tempPathField = typeof(FastSharedStringReader).GetField("_tempPath",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                tempFilePath = tempPathField?.GetValue(reader) as string;

                // Act
                reader.Dispose();
            }

            // Assert
            if (tempFilePath != null)
            {
                Assert.False(File.Exists(tempFilePath), "Temp file should be deleted after dispose");
            }
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Dispose_CalledMultipleTimes_DoesNotThrow()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["Test"]);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act & Assert - should not throw
            reader.Dispose();
            reader.Dispose();
            reader.Dispose();
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_AfterDispose_ThrowsObjectDisposedException()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["Test"]);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            var reader = FastSharedStringReader.Create(sharedStringPart);
            reader.Dispose();

            // Act & Assert
            Assert.Throws<ObjectDisposedException>(() => reader.GetString(0));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Edge Cases

    [Fact]
    public void Create_WithSingleString_Works()
    {
        // Arrange
        var filePath = CreateTestFileWithSharedStrings(["OnlyOne"]);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Assert
            Assert.Equal(1, reader.Count);
            Assert.Equal("OnlyOne", reader.GetString(0));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void GetString_WithDuplicateStrings_ReturnsCorrectly()
    {
        // Arrange
        var strings = new[] { "Duplicate", "Unique", "Duplicate", "Another", "Duplicate" };
        var filePath = CreateTestFileWithSharedStrings(strings);

        try
        {
            using var document = SpreadsheetDocument.Open(filePath, false);
            var sharedStringPart = document.WorkbookPart!.SharedStringTablePart!;
            using var reader = FastSharedStringReader.Create(sharedStringPart);

            // Act & Assert
            Assert.Equal("Duplicate", reader.GetString(0));
            Assert.Equal("Unique", reader.GetString(1));
            Assert.Equal("Duplicate", reader.GetString(2));
            Assert.Equal("Another", reader.GetString(3));
            Assert.Equal("Duplicate", reader.GetString(4));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Helper Methods

    private static string CreateTestFileWithSharedStrings(string[] strings)
    {
        var filePath = Path.GetTempFileName() + ".xlsx";

        using (var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = workbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Sheet1"
            };
            sheets.Append(sheet);

            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            var sharedStringTable = new SharedStringTable();

            foreach (var str in strings)
            {
                sharedStringTable.AppendChild(new SharedStringItem(new Text(str)));
            }

            sharedStringPart.SharedStringTable = sharedStringTable;

            var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>()!;
            var row = new Row { RowIndex = 1 };

            for (int i = 0; i < strings.Length; i++)
            {
                var cell = new Cell
                {
                    CellReference = $"{GetColumnLetter((uint)(i + 1))}1",
                    DataType = CellValues.SharedString,
                    CellValue = new CellValue(i.ToString())
                };
                row.Append(cell);
            }

            sheetData.Append(row);

            workbookPart.Workbook.Save();
        }

        return filePath;
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

    #endregion
}
