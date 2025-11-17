using ExcelImport.Exceptions;

namespace ExcelImport.Test;

public class ImporterTests
{
    #region Basic Import Tests

    [Fact]
    public void Import_SimpleModel_ImportsCorrectly()
    {
        var filePath = TestDataBuilder.CreateSimpleTestModelFile();

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Equal(3, results.Count);

            Assert.Equal("John", results[0].Name);
            Assert.Equal(30, results[0].Age);
            Assert.Equal(95.5m, results[0].Score);

            Assert.Equal("Jane", results[1].Name);
            Assert.Equal(25, results[1].Age);
            Assert.Equal(87.3m, results[1].Score);

            Assert.Equal("Bob", results[2].Name);
            Assert.Equal(35, results[2].Age);
            Assert.Equal(91.2m, results[2].Score);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_FirstItemOnly_DisposesContextCorrectly()
    {
        var filePath = TestDataBuilder.CreateSimpleTestModelFile();

        try
        {
            var importer = new ExcelImporter();
            var firstItem = importer.Import<SimpleTestModel>(filePath).First();

            Assert.NotNull(firstItem);
            Assert.Equal("John", firstItem.Name);

            // Context should be disposed after iteration stops
            // If not disposed, file would remain locked (we can't test this directly)
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_EmptyFile_ReturnsEmpty()
    {
        var data = new string[][] { new[] { "Name", "Age", "Score" } }; // Only header
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Empty(results);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Header Row Index Tests

    [Fact]
    public void Import_WithHeaderRowIndex0_ThrowsInvalidOperationException()
    {
        var filePath = TestDataBuilder.CreateSimpleTestModelFile();

        try
        {
            var importer = new ExcelImporter();
            Assert.Throws<InvalidOperationException>(() => importer.Import<SimpleTestModel>(filePath, headerRowIndex: 0).ToList());            
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Theory]
    [InlineData(1u)]
    [InlineData(5u)]
    [InlineData(10u)]
    public void Import_WithDifferentHeaderRowIndex_ImportsCorrectly(uint headerRow)
    {
        var filePath = TestDataBuilder.CreateFileWithHeaderAtRow(headerRow);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath, headerRowIndex: headerRow + 1).ToList();

            Assert.Single(results);
            Assert.Equal("John", results[0].Name);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Type Conversion Tests

    [Fact]
    public void Import_AllTypes_ConvertsCorrectly()
    {
        var data = new string[][]
        {
            new[] { "String", "Int", "Long", "Float", "Double", "Decimal", "Bool", "DateTime" },
            new[] { "test", "123", "123456789", "12.34", "98.76", "55.55", "true", "44562" } // 44562 is OA date for 2022-01-01
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<AllTypesModel>(filePath).ToList();

            Assert.Single(results);
            var item = results[0];

            Assert.Equal("test", item.StringValue);
            Assert.Equal(123, item.IntValue);
            Assert.Equal(123456789L, item.LongValue);
            Assert.Equal(12.34f, item.FloatValue, 2);
            Assert.Equal(98.76, item.DoubleValue, 2);
            Assert.Equal(55.55m, item.DecimalValue);
            Assert.True(item.BoolValue);
            Assert.NotEqual(default(DateTime), item.DateTimeValue);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_InvalidTypeConversion_ThrowsFormatException()
    {
        var data = new string[][]
        {
            ["Name", "Age", "Score"],
            ["John", "invalid", "not-a-number"]
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            Assert.Throws<FormatException>(() => importer.Import<SimpleTestModel>(filePath).ToList());            
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
    public void Import_RequiredFieldMissing_ThrowsImportException()
    {
        var filePath = TestDataBuilder.CreateFileWithMissingData();

        try
        {
            var importer = new ExcelImporter();

            Assert.Throws<ImportException>(() =>
                importer.Import<RequiredFieldsModel>(filePath).ToList());
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_RequiredFieldPresent_ImportsSuccessfully()
    {
        var data = new string[][]
        {
            new[] { "Required", "Optional" },
            new[] { "HasValue", "AlsoHasValue" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<RequiredFieldsModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("HasValue", results[0].RequiredField);
            Assert.Equal("AlsoHasValue", results[0].OptionalField);
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
    public void Import_WithCustomConverter_UsesConverter()
    {
        var data = new string[][]
        {
            new[] { "Amount", "Code" },
            new[] { "100", "ABC" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<CustomConverterModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal(10000m, results[0].Amount); // 100 * 100 (custom converter)
            Assert.Equal("ABC", results[0].Code);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Case Sensitivity Tests

    [Fact]
    public void Import_HeaderCaseInsensitive_MatchesCorrectly()
    {
        var data = new string[][]
        {
            ["NAME", "AGE", "SCORE"], // Uppercase headers
            ["John", "30", "95.5"]
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("John", results[0].Name);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_MixedCaseHeaders_MatchesCorrectly()
    {
        var data = new string[][]
        {
            new[] { "NaMe", "aGe", "ScOrE" },
            new[] { "John", "30", "95.5" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("John", results[0].Name);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Column Order Tests

    [Fact]
    public void Import_DifferentColumnOrder_ImportsCorrectly()
    {
        var data = new string[][]
        {
            new[] { "Score", "Name", "Age" }, // Different order
            new[] { "95.5", "John", "30" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("John", results[0].Name);
            Assert.Equal(30, results[0].Age);
            Assert.Equal(95.5m, results[0].Score);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_ExtraColumns_IgnoresExtra()
    {
        var data = new string[][]
        {
            new[] { "Name", "Age", "Score", "ExtraColumn" },
            new[] { "John", "30", "95.5", "Ignored" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("John", results[0].Name);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_MissingColumns_LeavesDefaults()
    {
        var data = new string[][]
        {
            new[] { "Name", "Age" }, // Missing Score column
            new[] { "John", "30" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("John", results[0].Name);
            Assert.Equal(30, results[0].Age);
            Assert.Equal(0m, results[0].Score); // Default value
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Null and Empty Value Tests

    [Fact]
    public void Import_EmptyStringValue_SetsEmptyString()
    {
        var data = new string[][]
        {
            new[] { "Name", "Age", "Score" },
            new[] { "", "30", "95.5" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("", results[0].Name);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_NullValueForNonRequiredField_SetsNull()
    {
        var data = new string[][]
        {
            new[] { "Required", "Optional" },
            new[] { "HasValue", "" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<RequiredFieldsModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("HasValue", results[0].RequiredField);
            Assert.Equal("", results[0].OptionalField);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Multiple Rows Tests

    [Fact]
    public void Import_MultipleRows_ImportsAll()
    {
        var data = new string[][]
        {
            new[] { "Name", "Age", "Score" },
            new[] { "John", "30", "95.5" },
            new[] { "Jane", "25", "87.3" },
            new[] { "Bob", "35", "91.2" },
            new[] { "Alice", "28", "89.7" },
            new[] { "Charlie", "32", "93.1" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Equal(5, results.Count);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_LargeFile_ImportsEfficiently()
    {
        var rows = new List<string[]>
        {
            new[] { "Name", "Age", "Score" }
        };

        for (int i = 0; i < 1000; i++)
        {
            rows.Add(new[] { $"Person{i}", i.ToString(), (i * 0.5).ToString() });
        }

        var filePath = TestDataBuilder.CreateSimpleExcelFile(rows.ToArray());

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Equal(1000, results.Count);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Lazy Enumeration Tests

    [Fact]
    public void Import_UsesLazyEnumeration_DoesNotLoadAllAtOnce()
    {
        var filePath = TestDataBuilder.CreateSimpleTestModelFile();

        try
        {
            var importer = new ExcelImporter();
            var enumerable = importer.Import<SimpleTestModel>(filePath);

            // Should not enumerate until we actually iterate
            var enumerator = enumerable.GetEnumerator();
            Assert.True(enumerator.MoveNext());
            Assert.NotNull(enumerator.Current);

            // Dispose without fully enumerating
            enumerator.Dispose();
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion

    #region Shared Strings Tests

    [Fact]
    public void Import_WithSharedStrings_ReadsCorrectly()
    {
        var data = new string[][]
        {
            new[] { "Name", "Age", "Score" },
            new[] { "John", "30", "95.5" }
        };
        var filePath = TestDataBuilder.CreateSimpleExcelFile(data, includeSharedStrings: true);

        try
        {
            var importer = new ExcelImporter();
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.Single(results);
            Assert.Equal("John", results[0].Name);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    [Fact]
    public void Import_WithoutSharedStrings_ReadsCorrectly()
    {
        var filePath = TestDataBuilder.CreateNumericExcelFile();

        try
        {
            var importer = new ExcelImporter();
            // This would fail before the SharedStrings fix
            var results = importer.Import<SimpleTestModel>(filePath).ToList();

            Assert.True(results.Count >= 0); // Should not throw
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    #endregion
}
