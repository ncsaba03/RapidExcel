using System.ComponentModel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters;

internal class DefaultTypeConverter<T> : TypeConverter<T?, string>
{

    public override CellValues CellType => CellValues.String;

    public override T? Convert(string value)
    {
        var converter = TypeDescriptor.GetConverter(typeof(T));
        var convertedValue = (T?)converter.ConvertFromInvariantString(value);
        return convertedValue;
    }

    public override CellValue? ConvertToCellValue(T? value)
    {
        var converter = TypeDescriptor.GetConverter(typeof(T));
        var converted = converter.ConvertToInvariantString(value);
        return new CellValue(converted ?? string.Empty);
    }
}
