using System.Runtime.CompilerServices;
using BankImport.Converters;
using BankImport.Model;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Converters;

namespace ExcelImport.Benchmarks;

/// <summary>
/// Benchmark for other methods where AggressiveInlining was removed
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 2, warmupCount: 1)]
public class OtherMethodsBenchmark
{
    private string[] _transactionTypes = [];
    private BankTransaction[] _transactions = [];
    private const int Iterations = 10000;

    [GlobalSetup]
    public void Setup()
    {
        // Setup transaction type strings for converter benchmark
        _transactionTypes = new[]
        {
            "KÁRTYATRANZAKCIÓ",
            "ÁTUTALÁS",
            "EGYÉB TERHELÉS",
            "DÍJ, KAMAT",
            "JÖVEDELEM",
            "EGYÉB JÓVÁÍRÁS"
        };

        // Setup transactions for parser benchmark
        _transactions = new[]
        {
            new BankTransaction
            {
                Date = DateTime.Now,
                Amount = -1000,
                Currency = "HUF",
                TransactionType = TransactionType.BankFeeOrInterest,
                Description = ""
            },
            new BankTransaction
            {
                Date = DateTime.Now,
                Amount = 5000,
                Currency = "HUF",
                TransactionType = TransactionType.Income,
                Description = "Salary payment from company"
            },
            new BankTransaction
            {
                Date = DateTime.Now,
                Amount = -2500,
                Currency = "HUF",
                TransactionType = TransactionType.Transfer,
                Description = "Transfer to account 12345678-12345678"
            }
        };
    }

    #region TransactionTypeConverter Benchmarks

    /// <summary>
    /// Baseline: TransactionTypeConverter.Convert WITHOUT AggressiveInlining
    /// Tests 10,000 conversions (simulating import of 10K rows)
    /// </summary>
    [Benchmark(Baseline = true)]
    public void TransactionTypeConverter_WithoutInlining()
    {
        var converter = new TransactionTypeConverter();
        for (int i = 0; i < Iterations; i++)
        {
            var typeStr = _transactionTypes[i % _transactionTypes.Length];
            var result = converter.Convert(typeStr);
            _ = result; // Prevent dead code elimination
        }
    }

    /// <summary>
    /// Test: TransactionTypeConverter.Convert WITH AggressiveInlining
    /// </summary>
    [Benchmark]
    public void TransactionTypeConverter_WithInlining()
    {
        var converter = new TransactionTypeConverterInlined();
        for (int i = 0; i < Iterations; i++)
        {
            var typeStr = _transactionTypes[i % _transactionTypes.Length];
            var result = converter.Convert(typeStr);
            _ = result;
        }
    }
       
    #endregion

    #region Helper Methods WITH AggressiveInlining

    /// <summary>
    /// Copy of TransactionTypeConverter WITH AggressiveInlining
    /// </summary>
    private class TransactionTypeConverterInlined : TypeConverter<TransactionType, string>
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

        public override CellValue? ConvertToCellValue(TransactionType value)
        {
            return new(value.ToString());
        }
    }

    #endregion
}
