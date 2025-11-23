using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters;

/// <summary>
/// Provides conversion between string representations and decimal values for Excel cell data, using invariant culture
/// formatting.
/// </summary>
/// <remarks>This converter is intended for use with Excel cells that store numeric amounts. It ensures that
/// decimal values are parsed and formatted consistently, regardless of locale, by using invariant culture. The
/// associated cell type is set to number, and the style index corresponds to the standard decimal format in
/// Excel.</remarks>
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
