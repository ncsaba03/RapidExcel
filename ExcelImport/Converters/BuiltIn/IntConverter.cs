using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn;

internal class IntConverter: TypeConverter<int, string>
{
    public override CellValues CellType => CellValues.Number;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override int Convert(string value)
    {
        return int.Parse(value);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(int value)
    {
        return new CellValue(value);
    }
}    