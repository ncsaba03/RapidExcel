using System.Reflection;
using ExcelImport.Converters;

public class PropertyImportInfo
{
    public required PropertyInfo Property { get; set; }
    public required string ColumnIdentifier { get; set; }
    public required TypeConverter TypeConverter { get; set; }    
    public int Position { get; set; }
    public bool Required { get; set; }
}
