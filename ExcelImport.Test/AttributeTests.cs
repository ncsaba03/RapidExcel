using ExcelImport.Attributes;
using ExcelImport.Converters;
using ExcelImport.Utils;

namespace ExcelImport.Test;

public class AttributeTests
{
    #region ExcelColumnAttribute Tests

    [Fact]
    public void ExcelColumnAttribute_Constructor_SetsName()
    {
        var attr = new ExcelColumnAttribute("TestColumn");

        Assert.Equal("TestColumn", attr.Name);
    }

    [Fact]
    public void ExcelColumnAttribute_Constructor_SetsPosition()
    {
        var attr = new ExcelColumnAttribute("TestColumn", position: 5);

        Assert.Equal(5, attr.Position);
    }

    [Fact]
    public void ExcelColumnAttribute_Constructor_SetsRequired()
    {
        var attr = new ExcelColumnAttribute("TestColumn", required: true);

        Assert.True(attr.Required);
    }

    [Fact]
    public void ExcelColumnAttribute_Constructor_SetsConverterType()
    {
        var attr = new ExcelColumnAttribute("TestColumn", typeConverter: typeof(TestCustomConverter));

        Assert.Equal(typeof(TestCustomConverter), attr.ConverterType);
    }

    [Fact]
    public void ExcelColumnAttribute_Constructor_AllParameters_SetsAllProperties()
    {
        var attr = new ExcelColumnAttribute("TestColumn", position: 3, typeConverter: typeof(TestCustomConverter), required: true);

        Assert.Equal("TestColumn", attr.Name);
        Assert.Equal(3, attr.Position);
        Assert.Equal(typeof(TestCustomConverter), attr.ConverterType);
        Assert.True(attr.Required);
    }

    [Fact]
    public void ExcelColumnAttribute_DefaultPosition_IsMinusOne()
    {
        var attr = new ExcelColumnAttribute("TestColumn");

        Assert.Equal(-1, attr.Position);
    }

    [Fact]
    public void ExcelColumnAttribute_DefaultRequired_IsFalse()
    {
        var attr = new ExcelColumnAttribute("TestColumn");

        Assert.False(attr.Required);
    }

    [Fact]
    public void ExcelColumnAttribute_DefaultConverterType_IsNull()
    {
        var attr = new ExcelColumnAttribute("TestColumn");

        Assert.Null(attr.ConverterType);
    }

    [Fact]
    public void ExcelColumnAttribute_CreateConverter_WithConverterType_ReturnsConverter()
    {
        var attr = new ExcelColumnAttribute("TestColumn", typeConverter: typeof(TestCustomConverter));

        var converter = attr.CreateConverter();

        Assert.NotNull(converter);
        Assert.IsType<TestCustomConverter>(converter);
    }

    [Fact]
    public void ExcelColumnAttribute_CreateConverter_WithoutConverterType_ReturnsNull()
    {
        var attr = new ExcelColumnAttribute("TestColumn");

        var converter = attr.CreateConverter();

        Assert.Null(converter);
    }

    #endregion

    #region PropertyCache Tests

    [Fact]
    public void PropertyCache_GetCachedProperties_WithAttributedProperties_ReturnsProperties()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(SimpleTestModel));

        Assert.Equal(3, properties.Count);
        Assert.Contains(properties, p => p.ColumnIdentifier == "Name");
        Assert.Contains(properties, p => p.ColumnIdentifier == "Age");
        Assert.Contains(properties, p => p.ColumnIdentifier == "Score");
    }

    [Fact]
    public void PropertyCache_GetCachedProperties_WithNoAttributes_ReturnsEmptyList()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(NoAttributesModel));

        Assert.Empty(properties);
    }

    [Fact]
    public void PropertyCache_GetCachedProperties_WithMixedAttributes_ReturnsOnlyAttributed()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(MixedAttributesModel));

        Assert.Single(properties);
        Assert.Equal("WithAttribute", properties[0].ColumnIdentifier);
    }

    [Fact]
    public void PropertyCache_GetCachedProperties_OrdersByPosition()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(PositionOrderingModel));

        Assert.Equal(4, properties.Count);
        Assert.Equal("First", properties[0].ColumnIdentifier);
        Assert.Equal("Second", properties[1].ColumnIdentifier);
        Assert.Equal("Third", properties[2].ColumnIdentifier);
        Assert.Equal("Last", properties[3].ColumnIdentifier); // Position not set, should be last
    }

    [Fact]
    public void PropertyCache_GetCachedProperties_CachesResult()
    {
        var properties1 = PropertyCache.GetCachedProperties(typeof(SimpleTestModel));
        var properties2 = PropertyCache.GetCachedProperties(typeof(SimpleTestModel));

        Assert.Same(properties1, properties2);
    }

    [Fact]
    public void PropertyCache_GetCachedProperties_WithRequiredField_SetsRequired()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(RequiredFieldsModel));

        var requiredProp = properties.First(p => p.ColumnIdentifier == "Required");
        var optionalProp = properties.First(p => p.ColumnIdentifier == "Optional");

        Assert.True(requiredProp.Required);
        Assert.False(optionalProp.Required);
    }

    [Fact]
    public void PropertyCache_GetCachedProperties_WithCustomConverter_CreatesConverter()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(CustomConverterModel));

        var amountProp = properties.First(p => p.ColumnIdentifier == "Amount");

        Assert.NotNull(amountProp.TypeConverter);
        Assert.IsType<TestCustomConverter>(amountProp.TypeConverter);
    }

    [Fact]
    public void PropertyCache_GetCachedProperties_WithoutCustomConverter_UsesDefaultConverter()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(SimpleTestModel));

        var ageProp = properties.First(p => p.ColumnIdentifier == "Age");

        Assert.NotNull(ageProp.TypeConverter);
        // Should use IntConverter for int type
        Assert.Contains("IntConverter", ageProp.TypeConverter.GetType().Name);
    }

    [Fact]
    public void PropertyCache_GetCachedProperties_AllTypes_CreatesCorrectConverters()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(AllTypesModel));

        var converterTypes = properties.Select(p => p.TypeConverter.GetType().Name).ToList();

        Assert.Contains("StringConverter", converterTypes);
        Assert.Contains("IntConverter", converterTypes);
        Assert.Contains("LongConverter", converterTypes);
        Assert.Contains("FloatConverter", converterTypes);
        Assert.Contains("DoubleConverter", converterTypes);
        Assert.Contains("DecimalConverter", converterTypes);
        Assert.Contains("BoolConverter", converterTypes);
        Assert.Contains("DateTimeConverter", converterTypes);
    }

    #endregion

    #region PropertyImportInfo Tests

    [Fact]
    public void PropertyImportInfo_Properties_CanBeSet()
    {
        var propertyInfo = typeof(SimpleTestModel).GetProperty("Name")!;
        var converter = new Converters.BuiltIn.StringConverter();

        var info = new PropertyImportInfo
        {
            Property = propertyInfo,
            ColumnIdentifier = "TestColumn",
            TypeConverter = converter,
            Position = 1,
            Required = true
        };

        Assert.Equal(propertyInfo, info.Property);
        Assert.Equal("TestColumn", info.ColumnIdentifier);
        Assert.Equal(converter, info.TypeConverter);
        Assert.Equal(1, info.Position);
        Assert.True(info.Required);
    }

    #endregion
}
