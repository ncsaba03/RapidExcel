using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn
{
    internal class BoolConverter : TypeConverter<bool, string>
    {
        public override CellValues CellType => CellValues.Boolean;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public override bool Convert(string value)
        {
            return bool.Parse(value);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public override CellValue? ConvertToCellValue(bool value)
        {
            return new(value);
        }
    }    
}