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

### Real-World Performance (349,495 records)

Quick performance overview using `Stopwatch` timing on **all_time_best.xlsx** (music charts, large SST):

| Operation | Time | Rows/sec | Notes |
|-----------|------|----------|-------|
| **Import (cold start)** | 2.2s | **158,000** | First run, no JIT/cache warmup |
| **Import (warm)** | 1.6s | **214,000** | Second run, JIT optimized |
| **Export Multi-Sheet (cold)** | 2.7s | **129,000** | First export, cold start |
| **Export Single (warm)** | 1.6s | **222,000** | Second export, optimized |

*Cold start = first run without JIT/cache optimization, Warm = subsequent runs*

### Detailed BenchmarkDotNet Analysis

BenchmarkDotNet measured on **Intel i7-11700K @ 3.60GHz**, **.NET 10.0**, **Release build**:

```
BenchmarkDotNet v0.15.6, Windows 11 (10.0.26200.7462)
11th Gen Intel Core i7-11700K 3.60GHz, 1 CPU, 16 logical and 8 physical cores
.NET SDK 10.0.101
Runtime: .NET 10.0.1, X64 RyuJIT x86-64-v4
InvocationCount=1, LaunchCount=2, UnrollFactor=1, WarmupCount=1
```

| Method                       | Mean       | Error    | StdDev   | Gen0        | Gen1      | Gen2      | Allocated  |
|----------------------------- |-----------:|---------:|---------:|------------:|----------:|----------:|-----------:|
| ImportTest                   |   932.2 ms |  2.55 ms |  3.57 ms |  55000.0000 |         - |         - |  443.43 MB |
| ImportTestWithParsing        | 1,122.6 ms |  7.40 ms | 10.85 ms |  89000.0000 |         - |         - |  711.95 MB |
| ExportTestWithMulitpleSheets | 3,267.9 ms | 12.04 ms | 16.08 ms | 285000.0000 | 1000.0000 | 1000.0000 | 2657.92 MB |
| ExportTest                   | 3,363.1 ms | 18.00 ms | 31.04 ms | 286000.0000 | 1000.0000 | 1000.0000 |  2658.9 MB |

#### Performance Summary

| Operation | Rows/sec | Memory | Gen2 | Description |
|-----------|----------|--------|------|-------------|
| **Import Only** | **477,400** | 443 MB | 0 | Streaming Excel parsing with SST optimization |
| **Import + Parse** | **396,500** | 712 MB | 0 | Complex financial data extraction |
| **Export Multi-Sheet** | **136,200** | 2.66 GB | 1K | Monthly sheets organization |
| **Export Single** | **132,300** | 2.66 GB | 1K | Single large worksheet |

**Performance Improvements** (SST Optimization - commit 200ab52):
- Import: **3.7x faster** (3.4s → 0.93s), **77% less memory** (1.93 GB → 443 MB)
- Export: **1.3x faster** (4.3s → 3.3s), **5% less memory** (2.8 GB → 2.66 GB)

**Key Optimization:** Replaced OpenXmlReader with manual XmlReader + O(1) SST lookups with pre-indexed SST.
- Before: O(N) SST lookups with object allocation
- After: Memory-mapped SST index with streaming XML parsing

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
