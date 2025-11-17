using BenchmarkDotNet.Running;
using ExcelImport.Benchmarks;

BenchmarkRunner.Run<ExcelImportBenchmark>();

// Other benchmarks (comment/uncomment as needed):
// BenchmarkRunner.Run<SheetHelperBenchmark>();
// BenchmarkRunner.Run<OtherMethodsBenchmark>();