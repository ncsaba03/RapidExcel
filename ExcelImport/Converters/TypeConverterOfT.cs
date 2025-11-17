using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters;

/// <summary>
/// Base class for all converters.
/// </summary>
/// <typeparam name="TTarget">The target type to convert to</typeparam>
/// <typeparam name="TSource">The source type to convert from</typeparam>
public abstract class TypeConverter<TTarget, TSource> : TypeConverter, IConverter<TTarget, TSource>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.String;

    /// <summary>
    /// Represents the style index of the cell.
    /// </summary>
    public override uint? StyleIndex => null;

    /// <summary>
    /// Converts the value to the target type.
    /// </summary>
    /// <param name="value"></param>
    /// <param name="throwOnConvertError"></param>
    /// <returns></returns>
    /// <exception cref="InvalidCastException"></exception>
    public override object? Convert(object value, bool throwOnConvertError = false)
    {
        var converted = Convert((TSource)value);

        if (converted is null && throwOnConvertError)
        {
            throw new InvalidCastException($"Cannot convert {value} to {typeof(TTarget)}");
        }

        return converted;
    }

    /// <summary>
    /// Converts the value to CellValue for the Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <param name="throwOnConvertError"></param>
    /// <returns></returns>
    /// <exception cref="InvalidCastException"></exception>
    public override CellValue? ConvertToCellValue(object value, bool throwOnConvertError = false)
    {
        var converted = ConvertToCellValue((TTarget)value);

        if (converted is null && throwOnConvertError)
        {
            throw new InvalidCastException($"Cannot convert {value} to {typeof(string)}");
        }

        return converted;
    }

    /// <summary>
    /// Converts the value to the target type.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public abstract TTarget? Convert(TSource value);

    /// <summary>
    /// Converts the value to the string representation of the target type.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public abstract CellValue? ConvertToCellValue(TTarget value);
}   