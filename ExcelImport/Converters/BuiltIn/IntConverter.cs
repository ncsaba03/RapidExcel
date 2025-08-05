using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

/// <summary>
/// Converter for integer values.
/// </summary>
internal class IntConverter: TypeConverter<int, string>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.Number;

    /// <summary>
    /// Converts a string value to an integer.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override int Convert(string value)
    {
        return int.Parse(value);
    }

    /// <summary>
    /// Converts the integer value to a CellValue for Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(int value)
    {
        return new CellValue(value);
    }
}    