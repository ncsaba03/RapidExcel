using System.Buffers;
using System.Diagnostics;
using BankImport.Parsers;
using ExcelImport;

string filePath = "C:\\test\\test500k.xlsx";
string exportPath = "C:\\test\\testexport.xlsx";

if (args.Length > 0)
{
    filePath = args[0] ?? filePath;
    exportPath = args[1] ?? exportPath;
}


Stopwatch stopwatch = new Stopwatch();
stopwatch.Start();
int count = 0;
var importer = new ExcelImporter();
var transactions = new List<TransactionDetail>(445_000);

foreach (var item in importer.Import<BankTransaction>(filePath, 10))
{
    var detail = TransactionDetailBuilder.Parse(item);
    transactions.Add(detail);
    count++;
}

Console.WriteLine("Total imported data:{0} Elapsed: {1} ms", count, stopwatch.ElapsedMilliseconds);
Console.ReadKey();
stopwatch.Restart();

var toExport = transactions.OrderBy(t => t.Date).GroupBy(t => new { sheetName = t.Date.ToString("yyyy MMMM") })
    .Select(g => (g.Key.sheetName, g.ToList()))
    .ToList();
var exporter = new ExcelExporter();

exporter.ExportSheetsWithWriter(toExport, exportPath);
Console.WriteLine("Total exported data:{0} Elapsed: {1} ms", count, stopwatch.ElapsedMilliseconds);
Console.ReadKey();


