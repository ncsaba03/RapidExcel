using System.ComponentModel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters;

/// <summary>
/// Default type converter for generic types.
/// </summary>
/// <typeparam name="T"></typeparam>
internal class DefaultTypeConverter<T> : TypeConverter<T?, string>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.String;

    /// <summary>
    /// Converts a string value to the specified type T.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public override T? Convert(string value)
    {
        var converter = TypeDescriptor.GetConverter(typeof(T));
        var convertedValue = (T?)converter.ConvertFromInvariantString(value);
        return convertedValue;
    }

    /// <summary>
    /// Converts the specified value of type T to a CellValue for Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public override CellValue? ConvertToCellValue(T? value)
    {
        var converter = TypeDescriptor.GetConverter(typeof(T));
        var converted = converter.ConvertToInvariantString(value);
        return new CellValue(converted ?? string.Empty);
    }
}
