using System.Reflection;
using ExcelImport.Converters;

internal class PropertyImportInfo
{
    /// <summary>
    ///  PropertyInfo for a property to be imported from an Excel file.
    /// </summary>
    public required PropertyInfo Property { get; set; }

    /// <summary>
    /// The identifier of the column in the Excel file.
    /// </summary>
    public required string ColumnIdentifier { get; set; }

    /// <summary>
    /// The type converter used to convert the value from the Excel file to the property type.
    /// </summary>
    public required TypeConverter TypeConverter { get; set; }

    /// <summary>
    /// The position of the column in the Excel file.
    /// </summary>
    public int Position { get; set; }

    /// <summary>
    /// Indicates whether the property is required in the Excel file.
    /// </summary>
    public bool Required { get; set; }
}
