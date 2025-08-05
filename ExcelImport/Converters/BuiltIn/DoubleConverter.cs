using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

/// <summary>
/// Converter for double values.
/// </summary>
internal class DoubleConverter : TypeConverter<double, string>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.Number;

    /// <summary>
    /// Represents the style index for double format in Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override double Convert(string value)
    {
        return double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Converts the double value to a CellValue for Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(double value)
    {
        return new CellValue(value);
    }
}    