using System.Diagnostics;
using Example.Test;
using RapidExcel;

string filePath = "C:\\test\\all_time_best.xlsx";
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
var charts = new List<Charts>(445_000);

foreach (var item in importer.Import<Charts>(filePath, 1))
{
    charts.Add(item);
    count++;
}

Console.WriteLine("Total imported data:{0} Elapsed: {1} ms", count, stopwatch.ElapsedMilliseconds);
Console.ReadKey();
stopwatch.Restart();

var toExport = charts.OrderBy(t => t.DateOfChart).GroupBy(t => new { sheetName = t.DateOfChart.ToString("yyyy MMMM") })
    .Select(g => (g.Key.sheetName, g.ToList()))
    .ToList();
var exporter = new ExcelExporter();

exporter.ExportSheets(toExport, exportPath);
Console.WriteLine("Total exported data:{0} Elapsed: {1} ms", count, stopwatch.ElapsedMilliseconds);
Console.ReadKey();


