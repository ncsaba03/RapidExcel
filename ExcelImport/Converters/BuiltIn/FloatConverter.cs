using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

internal class FloatConverter : TypeConverter<float, string>
{
    public override CellValues CellType => CellValues.Number;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override float Convert(string value)
    {
        return float.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(float value)
    {
        return new CellValue(value);
    }
}    