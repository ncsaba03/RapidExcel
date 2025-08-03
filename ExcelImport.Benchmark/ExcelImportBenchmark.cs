using BankImport.Parsers;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

namespace ExcelImport.Benchmarks
{
    [MemoryDiagnoser]
    [SimpleJob(RuntimeMoniker.Net80, launchCount: 2, warmupCount: 1)]
    public class ExcelImportBenchmark
    {
        private readonly ExcelImporter _importer = new ExcelImporter();
        private readonly ExcelExporter _exporter = new ExcelExporter();
        private  List<TransactionDetail> _parsedData = [];
        private List<(string, List<TransactionDetail>)> _parsedDataMultipleSheets = [];
        private string exportSingleFileName = $"{Path.GetRandomFileName()}.xlsx";
        private string exportMultipleSheetsFilename = $"{Path.GetRandomFileName()}.xlsx";

        [GlobalSetup]
        public void Setup()
        {            
            var rawData = _importer.Import<BankTransaction>("C:\\test\\test500k.xlsx", 10);
            _parsedData = [.. rawData.Select(TransactionDetailBuilder.Parse)];

            // Prepare multi-sheet data by date grouping
            _parsedDataMultipleSheets = _parsedData
                .GroupBy(t => t.Date.ToString("yyyy MMMM"))
                .Select(g => (g.Key, g.ToList()))
                .ToList();
        }

        [IterationCleanup]
        public void IterationCleanUp()
        {
            if (File.Exists(exportSingleFileName))
            {
                File.Delete(exportSingleFileName);
            }

            if (File.Exists(exportMultipleSheetsFilename))
            {
                File.Delete(exportMultipleSheetsFilename);
            }
        }

        [Benchmark]
        public void ImportTest()
        {
            foreach (var item in _importer.Import<BankTransaction>("C:\\test\\test500k.xlsx", 10))
            {
                _ = item; 
            }
        }

        [Benchmark]
        public void ImportTestWithParsing()
        {
            foreach (var item in _importer.Import<BankTransaction>("C:\\test\\test500k.xlsx", 10))
            {
                var detail = TransactionDetailBuilder.Parse(item);
                _ = detail; 
            }
        }

        [Benchmark]
        public void ExportTestWithMulitpleSheets()
        {
            _exporter.ExportSheetsWithWriter(_parsedDataMultipleSheets, $"{Path.GetRandomFileName}.xlsx");
        }

        [Benchmark]
        public void ExportTest()
        {
            _exporter.ExportWithWriter(_parsedData, exportSingleFileName);
        }
    }
}
