# RapidExcel
![RapidExcel Logo](icon.png)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-10.0-purple.svg)](https://dotnet.microsoft.com/)
[![Build](https://github.com/ncsaba03/ExcelImporter/actions/workflows/publish-nuget.yml/badge.svg)](https://github.com/ncsaba03/ExcelImporter/actions)

A very lightweight Excel import/export library for .NET 8+, designed for processing large datasets efficiently using OpenXml and modern C# features.


##  Features

- **Custom type converters** with attribute-based configuration
- **Streaming processing** for memory-efficient large file handling
- **Easy to Use**: Intuitive attributes and type converters make Excel processing straightforward
  
##  Use Cases

- Financial data processing (bank statements, transactions)
- Large dataset imports/exports
- Enterprise-scale data processing

##  Installation
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

## 📊 Performance Benchmarks

### Real-World Performance (445,060 records)

Quick performance overview using `Stopwatch` timing:

| Operation | Time | Rows/sec | Description |
|-----------|------|----------|-------------|
| **Import** | 5.5s | **80,700** | Pure Excel reading |
| **Export** | 4.7s | **94,500** | Excel file generation |
| **Round-trip** | 10.2s | **43,600** | Complete import + export cycle |
### Detailed BenchmarkDotNet Analysis

BenchmarkDotNet measured on **Intel i7-11700K @ 3.60GHz**, **.NET 8.0**, **Release build**:

```
BenchmarkDotNet v0.15.2, Windows 11
11th Gen Intel(R) Core(TM) i7-11700K @ 3.60GHz (3.60 GHz)
Runtime=.NET 8.0  LaunchCount=2  WarmupCount=1  
```

| Method                       | Mean    | Error    | StdDev   | Median  | Gen0        | Gen1      | Gen2      | Allocated |
|----------------------------- |--------:|---------:|---------:|--------:|------------:|----------:|----------:|----------:|
| ImportTest                   | 3.449 s | 0.0117 s | 0.0160 s | 3.440 s | 247000.0000 | 1000.0000 |         - |   1.93 GB |
| ImportTestWithParsing        | 3.573 s | 0.0098 s | 0.0138 s | 3.578 s | 281000.0000 | 1000.0000 |         - |   2.19 GB |
| ExportTestWithMulitpleSheets | 4.301 s | 0.0199 s | 0.0299 s | 4.293 s | 312000.0000 | 1000.0000 | 1000.0000 |    2.8 GB |
| ExportTest                   | 4.398 s | 0.0436 s | 0.0596 s | 4.359 s | 312000.0000 | 1000.0000 | 1000.0000 |    2.8 GB |

#### Performance Summary

| Operation | Rows/sec | Memory | Gen2 | Description |
|-----------|----------|--------|------|-------------|
| **Import Only** | **129,000** | 1.93 GB | 0 | Zero-allocation parsing |
| **Import + Parse** | **124,600** | 2.19 GB | 0 | Complex financial extraction |
| **Export Multi-Sheet** | **103,500** | 2.80 GB | 1K | Monthly sheets organization |
| **Export Single** | **101,200** | 2.80 GB | 1K | Single large worksheet |

## Example

### Complex Transaction Parsing

The library includes sophisticated parsing examples for CIB bank account statement financial data.

```csharp
// Parses complex bank transaction descriptions
var detail = TransactionDetailBuilder.Parse(bankTransaction);

// Extracts structured data from multi-line descriptions:
// - Card numbers and timestamps
// - Merchant information and location
// - Transaction IDs and references
```

##  Requirements

- .NET 8.0+
- DocumentFormat.OpenXml 3.3.0+

##  Project Structure

```
ExcelImporter/
├── ExcelImport/              # Core library
│   ├── Converters/           # Type conversion system
│   ├── Spreadsheet/          # Excel-specific utilities
│   └── Utils/                # Helper utilities
├── BankImport/               # Financial data example
├── Example/                  # Usage demonstrations
├── ExcelImport.Benchmark/    # Project used for benchmark
└── ExcelImport.Test/         # Unit tests

```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---
**Built with ❤️** 
