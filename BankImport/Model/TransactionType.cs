namespace BankImport.Model;

/// <summary>
/// Enum representing the type of a bank transaction.
/// </summary>
public enum TransactionType
{
    /// <summary>
    /// Unknown transaction type.
    /// </summary>
    Unknown,
    /// <summary>
    /// Transaction type for card transactions.
    /// </summary>
    Card,
    /// <summary>
    /// Transaction type for transfers.
    /// </summary>
    Transfer,
    /// <summary>
    /// Transaction type for bank fees or interest.
    /// </summary>
    BankFeeOrInterest,
    /// <summary>
    /// Transaction type for income transactions.
    /// </summary>
    Income,
}