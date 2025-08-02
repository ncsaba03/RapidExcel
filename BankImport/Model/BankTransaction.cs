using BankImport.Converters;
using BankImport.Model;
using ExcelImport.Attributes;

public record BankTransaction
{    
    [ExcelColumn("DÁTUM")]
    public DateTime Date { get; set; }
    
    [ExcelColumn("TRANZAKCIÓTÍPUS", typeConverter:typeof(TransactionTypeConverter))]
    public TransactionType TransactionType { get; set; }
    
    [ExcelColumn("KÖZLEMÉNY")]
    public string Description { get; set; } = string.Empty;

    [ExcelColumn("ÖSSZEG")]
    public decimal Amount { get; set; }

    [ExcelColumn("DEVIZANEM")]
    public string Currency { get; set; } = string.Empty;
}