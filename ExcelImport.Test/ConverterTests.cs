using System.Globalization;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Converters;
using ExcelImport.Converters.BuiltIn;

namespace ExcelImport.Test;

public class ConverterTests
{
    #region TypeConverter Factory Tests

    [Fact]
    public void CreateConverter_WithIntConverter_ReturnsIntConverter()
    {
        var converter = TypeConverter.CreateConverter(typeof(IntConverter));

        Assert.NotNull(converter);
        Assert.IsType<IntConverter>(converter);
    }

    [Fact]
    public void CreateConverter_WithSameType_ReturnsCachedInstance()
    {
        var converter1 = TypeConverter.CreateConverter(typeof(IntConverter));
        var converter2 = TypeConverter.CreateConverter(typeof(IntConverter));

        Assert.Same(converter1, converter2);
    }      

    [Fact]
    public void CreateConverter_WithNoParameterlessConstructor_ThrowsArgumentException()
    {
        Assert.Throws<ArgumentException>(() =>
            TypeConverter.CreateConverter(typeof(ConverterWithoutParameterlessConstructor)));
    }

    [Fact]
    public void GetConverterOfType_WithInt_ReturnsIntConverter()
    {
        var converter = TypeConverter.GetConverterOfType(typeof(int));

        Assert.NotNull(converter);
        Assert.IsType<IntConverter>(converter);
    }

    [Fact]
    public void GetConverterOfType_WithNullableInt_ReturnsIntConverter()
    {
        var converter = TypeConverter.GetConverterOfType(typeof(int?));

        Assert.NotNull(converter);
        Assert.IsType<IntConverter>(converter);
    }

    [Fact]
    public void GetConverterOfType_WithCustomType_ReturnsDefaultTypeConverter()
    {
        var converter = TypeConverter.GetConverterOfType(typeof(CustomType));

        Assert.NotNull(converter);
        Assert.Contains("DefaultTypeConverter", converter.GetType().Name);
    }

    [Fact]
    public void GetConverterOfType_CachesConverters()
    {
        var converter1 = TypeConverter.GetConverterOfType(typeof(int));
        var converter2 = TypeConverter.GetConverterOfType(typeof(int));

        Assert.Same(converter1, converter2);
    }

    [Fact]
    public void CreateDefaultTypeConverter_CreatesGenericConverter()
    {
        var converter = TypeConverter.CreateDefaultTypeConverter(typeof(CustomType));

        Assert.NotNull(converter);
        Assert.Contains("DefaultTypeConverter", converter.GetType().Name);
    }

    #endregion

    #region IntConverter Tests

    [Theory]
    [InlineData("0", 0)]
    [InlineData("123", 123)]
    [InlineData("-456", -456)]
    [InlineData("2147483647", int.MaxValue)]
    [InlineData("-2147483648", int.MinValue)]
    public void IntConverter_Convert_ValidValues_ReturnsInt(string input, int expected)
    {
        var converter = new IntConverter();

        var result = converter.Convert(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("12.34")]
    [InlineData("")]
    public void IntConverter_Convert_InvalidValues_ThrowsFormatException(string input)
    {
        var converter = new IntConverter();

        Assert.Throws<FormatException>(() => converter.Convert(input));
    }

    [Fact]
    public void IntConverter_Convert_Overflow_ThrowsOverflowException()
    {
        var converter = new IntConverter();

        Assert.Throws<OverflowException>(() => converter.Convert("2147483648"));
    }

    [Theory]
    [InlineData(0, "0")]
    [InlineData(123, "123")]
    [InlineData(-456, "-456")]
    public void IntConverter_ConvertToCellValue_ReturnsCorrectCellValue(int input, string expected)
    {
        var converter = new IntConverter();

        var result = converter.ConvertToCellValue(input);

        Assert.NotNull(result);
        Assert.Equal(expected, result.Text);
    }

    [Fact]
    public void IntConverter_CellType_IsNumber()
    {
        var converter = new IntConverter();

        Assert.Equal(CellValues.Number, converter.CellType);
    }

    #endregion

    #region LongConverter Tests

    [Theory]
    [InlineData("0", 0L)]
    [InlineData("123456789", 123456789L)]
    [InlineData("-987654321", -987654321L)]
    [InlineData("9223372036854775807", long.MaxValue)]
    [InlineData("-9223372036854775808", long.MinValue)]
    public void LongConverter_Convert_ValidValues_ReturnsLong(string input, long expected)
    {
        var converter = new LongConverter();

        var result = converter.Convert(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("12.34")]
    [InlineData("")]
    public void LongConverter_Convert_InvalidValues_ThrowsFormatException(string input)
    {
        var converter = new LongConverter();

        Assert.Throws<FormatException>(() => converter.Convert(input));
    }

    [Fact]
    public void LongConverter_CellType_IsNumber()
    {
        var converter = new LongConverter();

        Assert.Equal(CellValues.Number, converter.CellType);
    }

    #endregion

    #region FloatConverter Tests

    [Theory]
    [InlineData("0", 0f)]
    [InlineData("123.45", 123.45f)]
    [InlineData("-678.9", -678.9f)]
    [InlineData("3.4028235E+38", float.MaxValue)]
    [InlineData("-3.4028235E+38", float.MinValue)]
    public void FloatConverter_Convert_ValidValues_ReturnsFloat(string input, float expected)
    {
        var converter = new FloatConverter();

        var result = converter.Convert(input);

        Assert.Equal(expected, (float)result!, 5);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    public void FloatConverter_Convert_InvalidValues_ThrowsFormatException(string input)
    {
        var converter = new FloatConverter();

        Assert.Throws<FormatException>(() => converter.Convert(input));
    }

    [Fact]
    public void FloatConverter_CellType_IsNumber()
    {
        var converter = new FloatConverter();

        Assert.Equal(CellValues.Number, converter.CellType);
    }

    #endregion

    #region DoubleConverter Tests

    [Theory]
    [InlineData("0", 0.0)]
    [InlineData("123.456789", 123.456789)]
    [InlineData("-987.654321", -987.654321)]
    [InlineData("1.7976931348623157E+308", double.MaxValue)]
    [InlineData("-1.7976931348623157E+308", double.MinValue)]
    public void DoubleConverter_Convert_ValidValues_ReturnsDouble(string input, double expected)
    {
        var converter = new DoubleConverter();

        var result = converter.Convert(input);

        Assert.Equal(expected, (double)result!, 10);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    public void DoubleConverter_Convert_InvalidValues_ThrowsFormatException(string input)
    {
        var converter = new DoubleConverter();

        Assert.Throws<FormatException>(() => converter.Convert(input));
    }

    [Fact]
    public void DoubleConverter_CellType_IsNumber()
    {
        var converter = new DoubleConverter();

        Assert.Equal(CellValues.Number, converter.CellType);
    }

    #endregion

    #region DecimalConverter Tests

    [Theory]
    [InlineData("0", "0")]
    [InlineData("123.45", "123.45")]
    [InlineData("-678.9", "-678.9")]
    [InlineData("79228162514264337593543950335", "79228162514264337593543950335")] // decimal.MaxValue
    public void DecimalConverter_Convert_ValidValues_ReturnsDecimal(string input, string expectedStr)
    {
        var converter = new DecimalConverter();
        var expected = decimal.Parse(expectedStr, System.Globalization.CultureInfo.InvariantCulture);

        var result = converter.Convert(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    [InlineData("12.34.56")]
    public void DecimalConverter_Convert_InvalidValues_ThrowsFormatException(string input)
    {
        var converter = new DecimalConverter();

        Assert.Throws<FormatException>(() => converter.Convert(input));
    }

    [Theory]
    [InlineData(0, "0")]
    [InlineData(123.45, "123.45")]
    [InlineData(-678.9, "-678.9")]
    public void DecimalConverter_ConvertToCellValue_ReturnsCorrectCellValue(double inputDouble, string expected)
    {
        var converter = new DecimalConverter();
        var input = (decimal)inputDouble;

        var result = converter.ConvertToCellValue(input);

        Assert.NotNull(result);
        Assert.Equal(expected, result.Text);
    }

    [Fact]
    public void DecimalConverter_CellType_IsNumber()
    {
        var converter = new DecimalConverter();

        Assert.Equal(CellValues.Number, converter.CellType);
    }

    #endregion

    #region BoolConverter Tests

    [Theory]
    [InlineData("true", true)]
    [InlineData("True", true)]
    [InlineData("TRUE", true)]
    [InlineData("false", false)]
    [InlineData("False", false)]
    [InlineData("FALSE", false)]
    public void BoolConverter_Convert_ValidValues_ReturnsBool(string input, bool expected)
    {
        var converter = new BoolConverter();

        var result = converter.Convert(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("yes")]
    [InlineData("no")]
    [InlineData("1")]
    [InlineData("0")]
    [InlineData("")]
    [InlineData("maybe")]
    public void BoolConverter_Convert_InvalidValues_ThrowsFormatException(string input)
    {
        var converter = new BoolConverter();

        Assert.Throws<FormatException>(() => converter.Convert(input));
    }

    [Theory]
    [InlineData(true, "true")]
    [InlineData(false, "false")]
    public void BoolConverter_ConvertToCellValue_ReturnsCorrectCellValue(bool input, string expected)
    {
        var converter = new BoolConverter();

        var result = converter.ConvertToCellValue(input);

        Assert.NotNull(result);
        Assert.Equal(expected, result.Text);
    }

    [Fact]
    public void BoolConverter_CellType_IsBoolean()
    {
        var converter = new BoolConverter();

        Assert.Equal(CellValues.Boolean, converter.CellType);
    }

    #endregion

    #region StringConverter Tests

    [Theory]
    [InlineData("hello", "hello")]
    [InlineData("", "")]
    [InlineData("  spaces  ", "  spaces  ")]
    [InlineData("123", "123")]
    public void StringConverter_Convert_ReturnsString(string input, string expected)
    {
        var converter = new StringConverter();

        var result = converter.Convert(input);

        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("hello", "hello")]
    [InlineData("", "")]
    [InlineData("123", "123")]
    public void StringConverter_ConvertToCellValue_ReturnsCorrectCellValue(string input, string expected)
    {
        var converter = new StringConverter();

        var result = converter.ConvertToCellValue(input);

        Assert.NotNull(result);
        Assert.Equal(expected, result.Text);
    }

    [Fact]
    public void StringConverter_CellType_IsString()
    {
        var converter = new StringConverter();

        Assert.Equal(CellValues.String, converter.CellType);
    }

    #endregion

    #region DateTimeConverter Tests

    [Fact]
    public void DateTimeConverter_Convert_ValidOADate_ReturnsDateTime()
    {
        var converter = new DateTimeConverter();
        var oaDate = DateTime.Now.ToOADate().ToString(CultureInfo.InvariantCulture);

        var result = converter.Convert(oaDate);

        Assert.NotEqual(default, result);
        Assert.IsType<DateTime>(result);
    }

    [Theory]
    [InlineData("44562")] // 2022-01-01
    [InlineData("0")] // 1899-12-30
    public void DateTimeConverter_Convert_ValidOADateString_ReturnsDateTime(string input)
    {
        var converter = new DateTimeConverter();

        var result = converter.Convert(input);

        Assert.NotEqual(default, result);
    }

    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    [InlineData("not-a-date")]
    public void DateTimeConverter_Convert_InvalidValues_ThrowsFormatException(string input)
    {
        var converter = new DateTimeConverter();

        Assert.Throws<FormatException>(() => converter.Convert(input));
    }

    [Fact]
    public void DateTimeConverter_ConvertToCellValue_ReturnsOADate()
    {
        var converter = new DateTimeConverter();
        var date = new DateTime(2022, 1, 1);

        var result = converter.ConvertToCellValue(date);

        Assert.NotNull(result);
        var oaDate = double.Parse(result.Text);
        Assert.Equal(date.ToOADate(), oaDate);
    }

    [Fact]
    public void DateTimeConverter_CellType_IsNumber()
    {
        var converter = new DateTimeConverter();

        Assert.Equal(CellValues.Number, converter.CellType);
    }

    [Fact]
    public void DateTimeConverter_StyleIndex_Is1()
    {
        var converter = new DateTimeConverter();

        Assert.Equal(1u, converter.StyleIndex);
    }

    [Fact]
    public void DateTimeConverter_RoundTrip_PreservesDate()
    {
        var converter = new DateTimeConverter();
        var originalDate = new DateTime(2023, 6, 15, 14, 30, 0);

        var cellValue = converter.ConvertToCellValue(originalDate);
        var convertedBack = converter.Convert(cellValue!.Text);

        Assert.NotEqual(default, convertedBack);
        Assert.Equal(originalDate, convertedBack!);
    }

    #endregion

    #region DefaultTypeConverter Tests

    [Fact]
    public void DefaultTypeConverter_Convert_WithEnum_ConvertsCorrectly()
    {
        var converter = TypeConverter.CreateDefaultTypeConverter(typeof(TestEnum));

        var result = converter.Convert("Value1", false);

        Assert.Equal(TestEnum.Value1, result);
    }

    [Fact]
    public void DefaultTypeConverter_ConvertToCellValue_WithEnum_ConvertsCorrectly()
    {
        var converter = TypeConverter.CreateDefaultTypeConverter(typeof(TestEnum));

        var result = converter.ConvertToCellValue(TestEnum.Value2, false);

        Assert.NotNull(result);
        Assert.Equal("Value2", result.Text);
    }

    [Fact]
    public void DefaultTypeConverter_CellType_IsString()
    {
        var converter = TypeConverter.CreateDefaultTypeConverter(typeof(TestEnum));

        Assert.Equal(CellValues.String, converter.CellType);
    }

    #endregion

    #region Error Handling Tests

    [Fact]
    public void TypeConverter_Convert_WithBuiltInAndThrowOnError_ThrowsFormatException()
    {
        var converter = new IntConverter();

        Assert.Throws<FormatException>(() =>
            converter.Convert("invalid", throwOnConvertError: true));
    }

    [Fact]
    public void TypeConverter_ConvertToCellValue_WithNullBulitInAndThrowOnError_ThrowsNullReferenceException()
    {
        var converter = new IntConverter();

        Assert.Throws<NullReferenceException>(() =>
            converter.ConvertToCellValue(null!, throwOnConvertError: true));
    }

    [Fact]
    public void TypeConverter_ConvertToCellValue_WithNullAndThrowOnError_ThrowsInvalidCastException()
    {
        var converter = new DefaultTypeConverter<FirstClass>();
        Assert.Throws<InvalidCastException>(() =>
        converter.ConvertToCellValue(new SecondClass(), throwOnConvertError: true));
    }

    #endregion

    #region Helper Classes for Tests

    private class ConverterWithoutParameterlessConstructor : TypeConverter
    {
        public ConverterWithoutParameterlessConstructor(string param) { }

        public override CellValues CellType => CellValues.String;
        public override uint? StyleIndex => null;
        public override object? Convert(object value, bool throwOnConvertError = false) => null;
        public override CellValue? ConvertToCellValue(object value, bool throwOnConvertError = false) => null;
    }

    private class CustomType
    {
        public string Value { get; set; } = string.Empty;
    }

    private enum TestEnum
    {
        Value1,
        Value2,
        Value3
    }

    #endregion
}
