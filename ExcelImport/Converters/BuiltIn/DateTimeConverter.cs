using System.Globalization;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

/// <summary>
/// Converter for DateTime values.
/// </summary>
internal class DateTimeConverter : TypeConverter<DateTime, string>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.Number;

    /// <summary>
    /// Represents the style index for date format in Excel.
    /// </summary>
    public override uint? StyleIndex => 1U; // Date format

    /// <summary>
    /// Converts the string value to a DateTime object.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override DateTime Convert(string value)
    {
        return DateTime.FromOADate(double.Parse(value, CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Converts the DateTime value to a CellValue for Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(DateTime value)
    {
        return new CellValue(value.ToOADate());
    }
}
