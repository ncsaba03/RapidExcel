using System.Globalization;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

internal class LongConverter : TypeConverter<long, string>
{
    public override CellValues CellType => CellValues.Number;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override long Convert(string value)
    {
        return long.Parse(value);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(long value)
    {
        return new CellValue(value.ToString(CultureInfo.InvariantCulture));
    }
}    