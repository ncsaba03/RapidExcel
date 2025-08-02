using System.Globalization;
using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

internal class DateTimeConverter : TypeConverter<DateTime, string>
{
    public override CellValues CellType => CellValues.Number;
    public override uint? StyleIndex => 1U; // Date format

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override DateTime Convert(string value)
    {
        return DateTime.FromOADate(double.Parse(value));
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(DateTime value)
    {
        return new CellValue(value.ToOADate());
    }
}
