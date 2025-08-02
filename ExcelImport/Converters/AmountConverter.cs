using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters
{
    public class AmountConverter : TypeConverter<decimal, string>
    {
        public override uint? StyleIndex => 2;

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
}
