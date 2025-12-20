using BenchmarkDotNet.Running;
using RapidExcel.Benchmarks;

BenchmarkRunner.Run<ExcelImportBenchmark>();

// Other benchmarks (comment/uncomment as needed):
// BenchmarkRunner.Run<SheetHelperBenchmark>();
// BenchmarkRunner.Run<OtherMethodsBenchmark>();