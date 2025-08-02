using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Converters;

namespace BankImport.Converters;

internal class CardNumberConverter : TypeConverter<string, string?>
{
    public override string Convert(string? value)
    {
        return "BKC";
    }

    public override CellValue? ConvertToCellValue(string value)
    {
        var converted = value switch
        {
            "BKC" => "BKC",
            _ => "BKC"
        };

        return new CellValue(converted ?? string.Empty);
    }
}