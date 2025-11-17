using ExcelImport.Exceptions;

namespace ExcelImport.Test;

public class ContextTests
{
    #region ExcelImportContext Tests

    [Fact]
    public void ExcelImportContext_Constructor_ValidFile_OpensSuccessfully()
    {
        var data = new string[][]
        {
            new[] { "Name", "Age" },
            new[] { "John", "30" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            using var context = new ExcelImportContext(filePath);

            Assert.NotNull(context);
            Assert.NotNull(context.WorkbookPart);
            Assert.NotNull(context.WorksheetPart);
            Assert.NotNull(context.SharedStrings);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExcelImportContext_Constructor_FileWithoutSharedStrings_OpensSuccessfully()
    {
        var filePath = TestDataBuilder.CreateNumericExcelFile();

        try
        {
            using var context = new ExcelImportContext(filePath);

            Assert.NotNull(context);
            Assert.NotNull(context.WorkbookPart);
            Assert.NotNull(context.WorksheetPart);
            // SharedStrings can be null for numeric-only files
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExcelImportContext_Constructor_NonExistentFile_ThrowsException()
    {
        var filePath = "non_existent_file.xlsx";

        Assert.Throws<FileNotFoundException>(() => new ExcelImportContext(filePath));
    }

    [Fact]
    public void ExcelImportContext_Constructor_FileWithoutWorksheets_ThrowsInvalidOperationException()
    {
        var filePath = TestDataBuilder.CreateInvalidExcelFile();

        try
        {
            Assert.Throws<InvalidOperationException>(() => new ExcelImportContext(filePath));
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExcelImportContext_WorkbookPart_IsAccessible()
    {
        var data = new string[][] { new[] { "Test" } };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            using var context = new ExcelImportContext(filePath);

            var workbookPart = context.WorkbookPart;

            Assert.NotNull(workbookPart);
            Assert.NotNull(workbookPart.Workbook);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExcelImportContext_WorksheetPart_IsAccessible()
    {
        var data = new string[][] { new[] { "Test" } };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            using var context = new ExcelImportContext(filePath);

            var worksheetPart = context.WorksheetPart;

            Assert.NotNull(worksheetPart);
            Assert.NotNull(worksheetPart.Worksheet);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExcelImportContext_SharedStrings_IsAccessible()
    {
        var data = new string[][]
        {
            new[] { "Name", "Value" },
            new[] { "Test", "Data" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data, includeSharedStrings: true);

        try
        {
            using var context = new ExcelImportContext(filePath);

            var sharedStrings = context.SharedStrings;

            Assert.NotNull(sharedStrings);
            Assert.True(sharedStrings.ChildElements.Count > 0);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExcelImportContext_SheetRanges_ParsesCorrectly()
    {
        var data = new string[][]
        {
            new[] { "A", "B", "C" },
            new[] { "1", "2", "3" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            using var context = new ExcelImportContext(filePath);

            var sheetRanges = context.SheetRanges;

            Assert.NotNull(sheetRanges);
            Assert.Equal("A1", sheetRanges.Value.From.ToString());
            Assert.Equal("C2", sheetRanges.Value.To.ToString());
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExcelImportContext_Dispose_ReleasesResources()
    {
        var data = new string[][] { new[] { "Test" } };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var context = new ExcelImportContext(filePath);
            context.Dispose();

            // Should not throw after dispose
            context.Dispose(); // Double dispose should be safe
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void ExcelImportContext_UsingStatement_DisposesAutomatically()
    {
        var data = new string[][] { new[] { "Test" } };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            ExcelImportContext context;
            using (context = new ExcelImportContext(filePath))
            {
                Assert.NotNull(context.WorkbookPart);
            }

            // After using block, should be disposed
            // (We can't directly verify this without accessing private fields)
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region ImportException Tests

    [Fact]
    public void ImportException_Constructor_WithMessage_SetsProperties()
    {
        var identifier = "A1";
        var message = "Test error";

        var exception = new ImportException(identifier, message);

        Assert.Equal(identifier, exception.ImportIdentifier);
        Assert.Contains(identifier, exception.Message);
        Assert.Contains(message, exception.Message);
    }

    [Fact]
    public void ImportException_Constructor_WithInnerException_SetsProperties()
    {
        var identifier = "B2";
        var innerException = new InvalidOperationException("Inner error");

        var exception = new ImportException(identifier, innerException);

        Assert.Equal(identifier, exception.ImportIdentifier);
        Assert.Contains(identifier, exception.Message);
        Assert.Same(innerException, exception.InnerException);
    }

    [Fact]
    public void ImportException_Message_ContainsIdentifier()
    {
        var identifier = "C3";
        var message = "Validation failed";

        var exception = new ImportException(identifier, message);

        Assert.Contains($"[{identifier}]", exception.Message);
    }

    [Fact]
    public void ImportException_Message_ContainsErrorMessage()
    {
        var identifier = "D4";
        var message = "Required field is missing";

        var exception = new ImportException(identifier, message);

        Assert.Contains(message, exception.Message);
    }

    [Fact]
    public void ImportException_InheritsFromApplicationException()
    {
        var exception = new ImportException("E5", "Test");

        Assert.IsAssignableFrom<ApplicationException>(exception);
    }

    [Fact]
    public void ImportException_CanBeCaught()
    {
        var identifier = "F6";
        var message = "Test error";

        try
        {
            throw new ImportException(identifier, message);
        }
        catch (ImportException ex)
        {
            Assert.Equal(identifier, ex.ImportIdentifier);
            Assert.Contains(message, ex.Message);
        }
    }

    [Fact]
    public void ImportException_WithInnerException_PreservesInnerMessage()
    {
        var identifier = "G7";
        var innerMessage = "Inner exception message";
        var innerException = new FormatException(innerMessage);

        var exception = new ImportException(identifier, innerException);

        Assert.Contains(innerMessage, exception.Message);
        Assert.NotNull(exception.InnerException);
        Assert.Equal(innerMessage, exception.InnerException.Message);
    }

    #endregion
}
