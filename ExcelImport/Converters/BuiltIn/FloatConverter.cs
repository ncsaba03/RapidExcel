using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

/// <summary>
/// Converter for float values.
/// </summary>
internal class FloatConverter : TypeConverter<float, string>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.Number;

    /// <summary>
    /// Converts a string value to a float.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override float Convert(string value)
    {
        return float.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Converts the float value to a CellValue for Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(float value)
    {
        return new CellValue(value);
    }
}    