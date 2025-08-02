using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

internal class DecimalConverter : TypeConverter<decimal, string>
{
    public override CellValues CellType => CellValues.Number;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override decimal Convert(string value)
    {
        return decimal.Parse(value.AsSpan(), System.Globalization.CultureInfo.InvariantCulture);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(decimal value)
    {
        return new CellValue(value);
    }
}