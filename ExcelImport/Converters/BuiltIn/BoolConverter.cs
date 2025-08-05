using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

internal class BoolConverter : TypeConverter<bool, string>
{
    /// <summary>
    /// Represents the type of the cell
    /// </summary>
    public override CellValues CellType => CellValues.Boolean;

    /// <summary>
    /// Converts the string value to a boolean.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override bool Convert(string value)
    {
        return bool.Parse(value);
    }

    /// <summary>
    /// Converts the boolean value to a CellValue.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(bool value)
    {
        return new(value);
    }
}