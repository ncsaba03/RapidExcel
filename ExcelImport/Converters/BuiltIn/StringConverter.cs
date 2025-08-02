using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters.BuiltIn
{
    internal class StringConverter : TypeConverter<string, string>
    {
        public override DocumentFormat.OpenXml.Spreadsheet.CellValues CellType => DocumentFormat.OpenXml.Spreadsheet.CellValues.String;
        public override string Convert(string value)
        {
            return value;
        }
        public override CellValue? ConvertToCellValue(string value)
        {
            return new CellValue(value);
        }    
    }
}
