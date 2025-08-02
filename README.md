# ExcelImporter

A high-performance, zero-allocation Excel import/export library for .NET 8, designed for processing large datasets efficiently using modern C# features.

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

Tested on Intel i7-11700K, .NET 8 Release build:

| Operation | Rows | Time | Rows/sec |
|-----------|------|------|----------|
| Import | 445,060 | 5.5s | ~80,400 |
| Export | 445,060 | 4.9s | ~90,800 |
| **Total** | **445,060** | **10.4s** | **~85,600** |

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