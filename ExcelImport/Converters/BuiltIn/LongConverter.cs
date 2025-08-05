using System.Globalization;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

/// <summary>
/// Converter for long values.
/// </summary>
internal class LongConverter : TypeConverter<long, string>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.Number;

    /// <summary>
    /// Converts a string value to a long.        
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override long Convert(string value)
    {
        return long.Parse(value);
    }

    /// <summary>
    /// Converts the long value to a CellValue for Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(long value)
    {
        return new CellValue(value.ToString(CultureInfo.InvariantCulture));
    }
}    