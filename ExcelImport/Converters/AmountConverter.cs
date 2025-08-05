using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters;

public class AmountConverter : TypeConverter<decimal, string>
{
    /// <summary>
    /// Represents the style index for decimal format in Excel.
    /// </summary>
    public override uint? StyleIndex => 2;

    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.Number;

    /// <summary>
    /// Converts a string value to a decimal.
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
