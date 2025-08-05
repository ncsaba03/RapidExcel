namespace BankImport.Model;

public static class TransactionTypeExtensions
{
    /// <summary>
    /// Returns a bool value indicating if the transaction type is an income transaction.
    /// </summary>
    /// <param name="transactionType"></param>
    /// <returns></returns>
    public static bool IsIncome(this TransactionType transactionType)
    {
        return transactionType == TransactionType.Income;
    }
}


