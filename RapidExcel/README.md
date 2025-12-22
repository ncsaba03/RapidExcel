# RapidExcel
A very lightweight Excel import/export library for .NET 8, designed for processing large datasets efficiently using OpenXml and modern C# features.

## 📦 Installation
```bash
# Install via NuGet Package Manager
dotnet add package RapidExcel

# Or via Package Manager Console (Visual Studio)
Install-Package RapidExcel
```

##  Quick Start

### Define Your Model 

```csharp
//for import
public class BankTransaction
{
    [ExcelColumn("Date of Transaction")]
    public DateTime Date { get; set; }
    
    [ExcelColumn("TRANSTYPE", typeConverter: typeof(TransactionTypeConverter))]
    public TransactionType TransactionType { get; set; }
    
    [ExcelColumn("DESCR")]
    public string Description { get; set; } = string.Empty;

    [ExcelColumn("AMOUNT")]
    public decimal Amount { get; set; }

    [ExcelColumn("CUR")]
    public string Currency { get; set; } = string.Empty;
}

//for export
public record TransactionDetail
{
    [ExcelColumn("Date of Transaction", position: 1)]
    public DateTime Date { get; set; }

    [ExcelColumn("Description", position: 3)]
    public string? Payee { get; set; }

    [ExcelColumn("KH", position: 2, typeConverter: typeof(CardNumberConverter))]
    public string? CardNumber { get; set; }

    public string? AccountNumber { get; set; }

    [ExcelColumn("Amount", typeConverter: typeof(AmountConverter), position: 4)]
    public decimal Amount { get; set; }

    [ExcelColumn("Deviza", position: 5)]
    public string Currency { get; set; } = null!;

    public string? TransactionId { get; set; }

    public string? PayerId { get; set; }

    public string? City { get; set; }

    [ExcelColumn("IsExpense", position: 6)]
    public bool IsExpense { get; set; }

    [ExcelColumn("IsIncome", position: 7)]
    public bool IsIncome { get; set; }
}
```

### Import Data

```csharp
var importer = new ExcelImporter();

foreach (var transaction in importer.Import<BankTransaction>("data.xlsx", headerRowIndex: 10))
{
    // Process each transaction
    Console.WriteLine($"{transaction.Date}: {transaction.Amount} {transaction.Currency}");
}
```

### Export Data

```csharp
var transactions = GetTransactions(); // Your data source
var exporter = new ExcelExporter();

exporter.Export(transactions, "output.xlsx");

// Or export multiple sheets
var sheetData = transactions
    .GroupBy(t => t.Date.ToString("yyyy MMMM"))
    .Select(g => (g.Key, g.ToList()))
    .ToList();

exporter.ExportSheetsWithWriter(sheetData, "monthly_report.xlsx");
```

##  Custom Type Converters

Create custom converters for complex data transformations:

```csharp
public class TransactionTypeConverter : TypeConverter<TransactionType, string>
{
    public override TransactionType Convert(string value)
    {
        return value switch
        {
            "BCARD" => TransactionType.Card,
            "TRANSFER" => TransactionType.Transfer,
            "FEE" => TransactionType.BankFeeOrInterest,
            _ => TransactionType.Unknown
        };
    }

    public override CellValue? ConvertToCellValue(TransactionType value)
    {
        return new CellValue(value.ToString());
    }
}
```