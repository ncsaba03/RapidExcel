using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

/// <summary>
/// Converter for string values.
/// </summary>
internal class StringConverter : TypeConverter<string, string>
{
    /// <summary>
    /// Represents the type of the cell.
    /// </summary>
    public override CellValues CellType => CellValues.String;

    /// <summary>
    /// Converts a string value to a string.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public override string Convert(string value)
    {
        return value;
    }

    /// <summary>
    /// Converts the string value to a CellValue for Excel.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public override CellValue? ConvertToCellValue(string value)
    {
        return new CellValue(value);
    }
}
