using System.Runtime.CompilerServices;
using BankImport.Model;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Converters;

namespace BankImport.Converters;

/// <summary>
/// Converter for transaction types.
/// </summary>
public class TransactionTypeConverter : TypeConverter<TransactionType, string>
{
    public override CellValues CellType => CellValues.String;

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override TransactionType Convert(string value)
    {
        return value switch
        {
            "KÁRTYATRANZAKCIÓ" => TransactionType.Card,
            "ÁTUTALÁS" => TransactionType.Transfer,
            "EGYÉB TERHELÉS" => TransactionType.Transfer,
            "DÍJ, KAMAT" => TransactionType.BankFeeOrInterest,
            "JÖVEDELEM" => TransactionType.Income,
            "EGYÉB JÓVÁÍRÁS" => TransactionType.Income,
            _ => throw new ArgumentOutOfRangeException(nameof(value), value, "Invalid transaction type")
        };
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public override CellValue? ConvertToCellValue(TransactionType value)
    {
        return new (value.ToString());        
    }
}
