using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

internal class DoubleConverter : TypeConverter<double, string>
{
    public override CellValues CellType => CellValues.Number;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override double Convert(string value)
    {
        return double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(double value)
    {
        return new CellValue(value);
    }
}    