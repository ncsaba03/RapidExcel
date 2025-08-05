using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

/// <summary>
/// Converter for decimal values.
/// </summary>
internal class DecimalConverter : TypeConverter<decimal, string>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.Number;

    /// <summary>
    /// Represents the style index for decimal format in Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override decimal Convert(string value)
    {
        return decimal.Parse(value.AsSpan(), System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Converts the decimal value to a CellValue for Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(decimal value)
    {
        return new CellValue(value);
    }
}