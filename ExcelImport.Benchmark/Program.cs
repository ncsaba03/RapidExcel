using BenchmarkDotNet.Running;
using ExcelImport.Benchmarks;

// Run full import/export benchmark to measure GC pressure and compare with README
BenchmarkRunner.Run<ExcelImportBenchmark>();

// Other benchmarks (comment/uncomment as needed):
// BenchmarkRunner.Run<SheetHelperBenchmark>();
// BenchmarkRunner.Run<OtherMethodsBenchmark>();