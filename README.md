# ExcelImporter

A high-performance lightweight Excel import/export library for .NET 8, designed for processing large datasets efficiently using modern C# features.

## 🚀 Performance

- **80,000+ rows/sec** import speed
- **90,000+ rows/sec** export speed  
- **445,000 rows** processed in ~10.4 seconds

## ✨ Features

- **Zero-allocation parsing** with `Span<T>` and `ReadOnlySpan<char>`
- **Custom type converters** with attribute-based configuration
- **Streaming processing** for memory-efficient large file handling
- **Thread-safe property caching** with reflection optimization
- **Custom enumerators** for allocation-free string operations
- **OpenXML streaming** for optimal performance

## 🎯 Use Cases

- Financial data processing (bank statements, transactions)
- Large dataset imports/exports
- Performance-critical applications
- Memory-constrained environments
- Enterprise-scale data processing

## 📦 Installation

```bash
# Clone the repository
git clone https://github.com/ncsaba03/ExcelImporter.git
cd ExcelImporter

# Build the solution
dotnet build -c Release
```

## 🔧 Quick Start

### Define Your Model

```csharp
public class BankTransaction
{
    [ExcelColumn("DÁTUM")]
    public DateTime Date { get; set; }
    
    [ExcelColumn("TRANZAKCIÓTÍPUS", typeConverter: typeof(TransactionTypeConverter))]
    public TransactionType TransactionType { get; set; }
    
    [ExcelColumn("KÖZLEMÉNY")]
    public string Description { get; set; } = string.Empty;

    [ExcelColumn("ÖSSZEG")]
    public decimal Amount { get; set; }

    [ExcelColumn("DEVIZANEM")]
    public string Currency { get; set; } = string.Empty;
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

## 🔄 Custom Type Converters

Create custom converters for complex data transformations:

```csharp
public class TransactionTypeConverter : TypeConverter<TransactionType, string>
{
    public override TransactionType Convert(string value)
    {
        return value switch
        {
            "KÁRTYATRANZAKCIÓ" => TransactionType.Card,
            "ÁTUTALÁS" => TransactionType.Transfer,
            "DÍJ, KAMAT" => TransactionType.BankFeeOrInterest,
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


## 🏗️ Architecture

### Zero-Allocation Parsing

```csharp
// Uses ReadOnlySpan<char> for zero-copy string operations
public static SheetCell Parse(ReadOnlySpan<char> cellString)
{
    var cell = cellString.Trim();
    var col = cell[..colIndex];
    int row = int.Parse(cell[colIndex..]);
    return new(col.ToString(), row);
}
```

### Custom Enumerators

```csharp
// SpanSplitEnumerator for allocation-free string splitting
public ref struct SpanSplitEnumerator
{
    private ReadOnlySpan<char> _span;
    private readonly char _separator;
    // Zero heap allocation string splitting
}
```

### Property Caching

```csharp
// Thread-safe reflection caching with expression trees
private static readonly ConcurrentDictionary<Type, List<PropertyImportInfo>> _propertyCache = new();

// Compiled factory methods for performance
private static Func<TypeConverter> CreateFactory(Type type)
{
    var ctor = type.GetConstructor(Type.EmptyTypes);
    var newExpr = Expression.New(ctor);
    return Expression.Lambda<Func<TypeConverter>>(newExpr).Compile();
}
```

## 🧪 Advanced Usage

### Complex Transaction Parsing

The library includes sophisticated parsing for financial data:

```csharp
// Parses complex bank transaction descriptions
var detail = TransactionDetailBuilder.Parse(bankTransaction);

// Extracts structured data from multi-line descriptions:
// - Card numbers and timestamps
// - Merchant information and location
// - Transaction IDs and references
```

### Memory-Efficient Processing

```csharp
// Stream processing for large files
using var context = new ExcelImportContext(filePath);
using var reader = OpenXmlReader.Create(context.WorksheetPart);

// Yield return for lazy enumeration
public IEnumerable<T> Import<T>(string filePath) where T : new()
{
    // Process row by row without loading entire file
    yield return item;
}
```

## 🛠️ Requirements

- .NET 8.0+
- DocumentFormat.OpenXml 3.3.0+

## 📁 Project Structure

```
ExcelImporter/
├── ExcelImport/              # Core library
│   ├── Converters/           # Type conversion system
│   ├── Spreadsheet/          # Excel-specific utilities
│   └── Utils/                # Helper utilities
├── BankImport/               # Financial data example
├── Example/                  # Usage demonstrations
└── ExcelImport.Test/         # Unit tests
```

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---
**Built with ❤️ for performance and efficiency**
