// See https://aka.ms/new-console-template for more information
using BenchmarkDotNet.Running;
using ExcelImport.Benchmarks;

BenchmarkRunner.Run<ExcelImportBenchmark>();