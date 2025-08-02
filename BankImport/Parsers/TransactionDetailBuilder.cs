using System.Collections.Frozen;
using System.Runtime.CompilerServices;
using BankImport.Model;
using ExcelImport.Utils;

namespace BankImport.Parsers;

/// <summary>
/// Builder for creating transaction details.
/// </summary>
public readonly ref struct TransactionDetailBuilder
{
    private readonly TransactionDetail _detail;

    /// <summary>
    /// Initializes a new instance of the <see cref="TransactionDetailBuilder"/> class.
    /// </summary>
    public TransactionDetailBuilder()
    {
        _detail = new TransactionDetail();
    }

    /// <summary>
    /// Tries to parse a <see cref="TransactionDetail"/> from a <see cref="BankTransaction"/>.
    /// </summary>
    /// <param name="transaction"></param>
    /// <param name="transactionDetail"></param>
    /// <returns></returns>
    public static bool TryParse(BankTransaction transaction, out TransactionDetail? transactionDetail)
    {
        transactionDetail = null;
        try
        {
            transactionDetail = Parse(transaction);
            return true;
        }
        catch (Exception ex) when (ex is FormatException || ex is ArgumentException)
        {
            return false;
        }
    }

    /// <summary>
    /// Parses a <see cref="TransactionDetail"/> from a <see cref="BankTransaction"/>.
    /// </summary>
    /// <param name="transaction"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    public static TransactionDetail Parse(BankTransaction transaction)
    {
        var details = transaction.Description.AsSpan();

        return transaction.TransactionType switch
        {
            TransactionType.Card => ParseCard(details, transaction),
            TransactionType.Transfer => ParseTransfer(details, transaction),
            TransactionType.BankFeeOrInterest => ParseBankFee(transaction),
            TransactionType.Income => ParseIncome(details, transaction),
            _ => throw new ArgumentOutOfRangeException(nameof(transaction.TransactionType), transaction.TransactionType, "Invalid transaction type")
        };
    }

    /// <summary>
    /// Parses a <see cref="TransactionDetail"/> for income transactions.
    /// </summary>
    /// <param name="details"></param>
    /// <param name="bankTransaction"></param>
    /// <returns></returns>
    /// <exception cref="FormatException"></exception>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static TransactionDetail ParseIncome(ReadOnlySpan<char> details, BankTransaction bankTransaction)
    {
        if (details.IsEmpty) throw new FormatException("Income data is missing");
        return ParseSimpleDescription(details, bankTransaction);
    }

    /// <summary>
    /// Parses a <see cref="TransactionDetail"/> for bank fees or interest.
    /// </summary>
    /// <param name="bankTransaction"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static TransactionDetail ParseBankFee(BankTransaction bankTransaction)
        => new TransactionDetail
        {
            Date = bankTransaction.Date,
            Amount = -1 * bankTransaction.Amount,
            Currency = bankTransaction.Currency,
            Payee = string.IsInterned("CIB BANK") ?? "CIB BANK",
            City = string.IsInterned("Budapest") ?? "Budapest",
            PayerId = string.IsInterned("CIB") ?? "CIB",
            IsExpense = bankTransaction.Amount < 0,
            IsIncome = false
        };

    /// <summary>
    /// Parses a <see cref="TransactionDetail"/> for transfer transactions.
    /// </summary>
    /// <param name="details"></param>
    /// <param name="bankTransaction"></param>
    /// <returns></returns>
    /// <exception cref="FormatException"></exception>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static TransactionDetail ParseTransfer(ReadOnlySpan<char> details, BankTransaction bankTransaction)
    {
        if (details.IsEmpty) throw new FormatException("Transfer data is missing");
        return ParseSimpleDescription(details, bankTransaction);
    }

    /// <summary>
    /// Parses a <see cref="TransactionDetail"/> for card transactions.
    /// </summary>
    /// <param name="details"></param>
    /// <returns></returns>
    /// <exception cref="FormatException"></exception>
    public static TransactionDetail ParseCard(ReadOnlySpan<char> details, BankTransaction transaction)
    {
        if (details.IsEmpty) throw new FormatException("Card data is missing");

        var builder = new TransactionDetailBuilder();
        var enumerator = details.EnumerateLines();

        if (!enumerator.MoveNext()) throw new FormatException("Card data and transaction date is missing");
        builder.SetCardAndDate(enumerator.Current);

        if (!enumerator.MoveNext()) throw new FormatException("Amount and currency is missing");
        builder.SetAmountAndCurrency(enumerator.Current,transaction);

        if (!enumerator.MoveNext()) throw new FormatException("No additional data");
        builder.SetAdditionalDataAndCity(enumerator.Current);

        if (!enumerator.MoveNext()) throw new FormatException("Payee information is missing");
        builder.SetPayeeAndIds(enumerator.Current);

        return builder._detail;
    }

    /// <summary>
    /// Parses a simple description for transactions that do not require special handling.
    /// </summary>
    /// <param name="details"></param>
    /// <param name="bankTransaction"></param>
    /// <returns></returns>
    private static TransactionDetail ParseSimpleDescription(ReadOnlySpan<char> details, BankTransaction bankTransaction)
    {
        var lineEnumerator = details.EnumerateLines();
        TransactionDetail detail = new TransactionDetail
        {
            Date = bankTransaction.Date,
            Amount = -1 * bankTransaction.Amount,
            Currency = bankTransaction.Currency,
            AccountNumber = lineEnumerator.MoveNext() ? lineEnumerator.Current.ToString() : null,
            Payee = lineEnumerator.MoveNext() ? lineEnumerator.Current.ToString() : "N/A",
            IsExpense = bankTransaction.Amount < 0,
            IsIncome = bankTransaction.TransactionType.IsIncome(),
        };

        var dictionary = new Dictionary<string, string?>(2, StringComparer.OrdinalIgnoreCase);

        uint index = 1;
        foreach (var item in lineEnumerator)
        {
            dictionary[$"Param{index++}"] = item.ToString();
        }

        detail.AdditionalData = dictionary;

        return detail;
    }

    /// <summary>
    /// Sets the card number and date from the given line.
    /// </summary>
    /// <param name="line"></param>
    private void SetCardAndDate(ReadOnlySpan<char> line)
    {
        _detail.CardNumber = string.IsInterned(line[..19].ToString()) ?? string.Intern(line[..19].ToString()) ;
        _detail.Date = DateTime.ParseExact(line[20..35], "yyyyMMdd HHmmss", null);
    }

    /// <summary>
    /// Sets the amount and currency from the given line.
    /// </summary>
    /// <param name="line"></param>
    /// <exception cref="ArgumentException"></exception>
    private void SetAmountAndCurrency(ReadOnlySpan<char> line ,BankTransaction transaction)
    {
        
        _detail.Currency = transaction.Currency;
        _detail.Amount = -1 * transaction.Amount;        
        _detail.IsExpense = transaction.Amount < 0;
        _detail.IsIncome = false;
    }

    /// <summary>
    /// Sets the additional data and city from the given line.
    /// </summary>
    /// <param name="line"></param>
    private void SetAdditionalDataAndCity(ReadOnlySpan<char> line)
    {
        var additionalData = new Dictionary<string, string?>(2, StringComparer.OrdinalIgnoreCase);
        var split = new SpanSplitEnumerator(line, ' ');

        additionalData["Param1"] = split.MoveNext() ? split.Current.ToString() : null;
        additionalData["Param2"] = split.MoveNext() ? split.Current.ToString() : null;

        var city = line[split.Position..].Trim();
        _detail.City = city.IsEmpty ? null : city.ToString();
        _detail.AdditionalData = additionalData;
    }

    /// <summary>
    /// Sets the payee and IDs from the given line.
    /// </summary>
    /// <param name="line"></param>
    private void SetPayeeAndIds(ReadOnlySpan<char> line)
    {
        _detail.Payee = line[..18].TrimEnd().ToString();

        var payerId = line[19..28].Trim();
        var transactionId = line[29..].Trim();

        _detail.PayerId = payerId.IsEmpty ? null : payerId.ToString();
        _detail.TransactionId = transactionId.IsEmpty ? null : transactionId.ToString();
    }
}
