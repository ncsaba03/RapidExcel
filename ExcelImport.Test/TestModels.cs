using ExcelImport.Attributes;
using ExcelImport.Converters;

namespace ExcelImport.Test;

/// <summary>
/// Test models used across multiple test files
/// </summary>

public class SimpleTestModel
{
    [ExcelColumn("Name", position: 1)]
    public string Name { get; set; } = string.Empty;

    [ExcelColumn("Age", position: 2)]
    public int Age { get; set; }

    [ExcelColumn("Score", position: 3)]
    public decimal Score { get; set; }
}

public class RequiredFieldsModel
{
    [ExcelColumn("Required", required: true)]
    public string RequiredField { get; set; } = string.Empty;

    [ExcelColumn("Optional")]
    public string? OptionalField { get; set; }
}

public class CustomConverterModel
{
    [ExcelColumn("Amount", typeConverter: typeof(TestCustomConverter))]
    public decimal Amount { get; set; }

    [ExcelColumn("Code")]
    public string Code { get; set; } = string.Empty;
}

public class AllTypesModel
{
    [ExcelColumn("String")]
    public string StringValue { get; set; } = string.Empty;

    [ExcelColumn("Int")]
    public int IntValue { get; set; }

    [ExcelColumn("Long")]
    public long LongValue { get; set; }

    [ExcelColumn("Float")]
    public float FloatValue { get; set; }

    [ExcelColumn("Double")]
    public double DoubleValue { get; set; }

    [ExcelColumn("Decimal")]
    public decimal DecimalValue { get; set; }

    [ExcelColumn("Bool")]
    public bool BoolValue { get; set; }

    [ExcelColumn("DateTime")]
    public DateTime DateTimeValue { get; set; }
}

public class PositionOrderingModel
{
    [ExcelColumn("Third", position: 3)]
    public string Third { get; set; } = string.Empty;

    [ExcelColumn("First", position: 1)]
    public string First { get; set; } = string.Empty;

    [ExcelColumn("Second", position: 2)]
    public string Second { get; set; } = string.Empty;

    [ExcelColumn("Last", position: 4)]
    public string Last { get; set; } = string.Empty;
}

public class NoAttributesModel
{
    public string Name { get; set; } = string.Empty;
    public int Age { get; set; }
}

public class MixedAttributesModel
{
    [ExcelColumn("WithAttribute")]
    public string WithAttribute { get; set; } = string.Empty;

    public string WithoutAttribute { get; set; } = string.Empty;
}

public class FirstClass
{
    [ExcelColumn("Value")]
    public string Value { get; set; } = string.Empty;
        
}

public class SecondClass
{
    public decimal Amount { get; set; }
}

/// <summary>
/// Test custom converter for testing purposes
/// </summary>
public class TestCustomConverter : TypeConverter<decimal, string>
{
    public override decimal Convert(string value)
    {
        return decimal.Parse(value) * 100;
    }

    public override DocumentFormat.OpenXml.Spreadsheet.CellValue? ConvertToCellValue(decimal value)
    {
        return new DocumentFormat.OpenXml.Spreadsheet.CellValue((value / 100).ToString());
    }
}
