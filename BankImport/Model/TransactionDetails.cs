using BankImport.Converters;
using ExcelImport.Attributes;
using ExcelImport.Converters;

namespace BankImport.Model;

public record TransactionDetail
{
    /// <summary>
    /// Date of the transaction.
    /// </summary>
    [ExcelColumn("Dátum", position: 1)]
    public DateTime Date { get; set; }

    /// <summary>
    /// Payee or sender of the transaction.
    /// </summary>
    [ExcelColumn("Megnevezés", position: 3)]
    public string Payee { get; set; } = null!;

    /// <summary>
    /// Card number
    /// </summary>    
    [ExcelColumn("KH", position: 2, typeConverter: typeof(CardNumberConverter))]
    public string? CardNumber { get; set; }

    /// <summary>
    /// Account number
    /// </summary>
    public string? AccountNumber { get; set; }

    /// <summary>
    /// Amount of the transaction.
    /// </summary>
    [ExcelColumn("Összeg", typeConverter: typeof(AmountConverter), position: 4)]
    public decimal Amount { get; set; }

    /// <summary>
    /// Additional data related to the transaction.
    /// </summary>
    public Dictionary<string, string?>? AdditionalData { get; set; }

    /// <summary>
    /// Currency of the transaction.
    /// </summary>
    [ExcelColumn("Deviza", position: 5)]
    public string Currency { get; set; } = null!;

    /// <summary>
    /// Transaction identifier or reference number.
    /// </summary>
    public string? TransactionId { get; set; }

    /// <summary>
    /// Payer reference number.
    /// </summary>
    public string? PayerId { get; set; }

    /// <summary>
    /// City where the transaction took place.
    /// </summary>
    public string? City { get; set; }

    /// <summary>
    /// Indicates if the transaction is an expense.
    /// </summary>
    [ExcelColumn("Kiadás", position: 6)]
    public bool IsExpense { get; set; }

    /// <summary>
    /// Indicates if the transaction is an income.
    /// </summary>
    [ExcelColumn("Jövedelem", position: 7)]
    public bool IsIncome { get; set; }
}