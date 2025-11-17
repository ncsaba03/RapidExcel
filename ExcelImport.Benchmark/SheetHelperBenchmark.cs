using System.Runtime.CompilerServices;
using BenchmarkDotNet.Attributes;
using ExcelImport.Spreadsheet;

namespace ExcelImport.Benchmarks;

/// <summary>
/// Benchmark to measure the impact of AggressiveInlining on TransformToCharacterIndex
/// This simulates realistic export scenarios where the method is called for every cell.
/// </summary>
[MemoryDiagnoser]
[SimpleJob(launchCount: 2, warmupCount: 1)]
public class SheetHelperBenchmark
{
    private uint[] _columnIndices = [];
    private const int Rows = 1000;
    private const int Columns = 10;

    [GlobalSetup]
    public void Setup()
    {
        // Prepare column indices similar to export scenario (1-based indices for columns A-J)
        _columnIndices = Enumerable.Range(1, Columns).Select(i => (uint)i).ToArray();
    }

    /// <summary>
    /// Baseline: Current implementation WITHOUT AggressiveInlining
    /// Simulates export of 1000 rows × 10 columns = 10,000 calls
    /// </summary>
    [Benchmark(Baseline = true)]
    public void TransformToCharacterIndex_WithoutInlining()
    {
        for (int row = 0; row < Rows; row++)
        {
            for (int col = 0; col < _columnIndices.Length; col++)
            {
                var cellRef = SheetHelper.TransformToCharacterIndex(_columnIndices[col]);
                // Simulate using the result (prevent dead code elimination)
                _ = cellRef;
            }
        }
    }

    /// <summary>
    /// Test: Implementation WITH AggressiveInlining
    /// Simulates export of 1000 rows × 10 columns = 10,000 calls
    /// </summary>
    [Benchmark]
    public void TransformToCharacterIndex_WithInlining()
    {
        for (int row = 0; row < Rows; row++)
        {
            for (int col = 0; col < _columnIndices.Length; col++)
            {
                var cellRef = TransformToCharacterIndexInlined(_columnIndices[col]);
                // Simulate using the result (prevent dead code elimination)
                _ = cellRef;
            }
        }
    }

    /// <summary>
    /// Copy of TransformToCharacterIndex WITH AggressiveInlining attribute for comparison
    /// </summary>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static string TransformToCharacterIndexInlined(uint colIndex)
    {
        if (colIndex == 0)
            throw new ArgumentOutOfRangeException(nameof(colIndex), "Column index must be >= 1");

        Span<char> buffer = stackalloc char[8];
        int pos = buffer.Length;

        uint index = colIndex;
        while (index > 0)
        {
            index--;
            buffer[--pos] = (char)('A' + (index % 26));
            index /= 26;
        }

        return new(buffer[pos..]);
    }

    /// <summary>
    /// Additional test: Larger column indices (testing AA, AAA, etc.)
    /// Tests columns 1-100 to include multi-character column names
    /// </summary>
    [Benchmark]
    public void TransformToCharacterIndex_LargeIndices_WithoutInlining()
    {
        for (int row = 0; row < 100; row++)
        {
            for (uint col = 1; col <= 100; col++)
            {
                var cellRef = SheetHelper.TransformToCharacterIndex(col);
                _ = cellRef;
            }
        }
    }

    /// <summary>
    /// Additional test: Larger column indices WITH AggressiveInlining
    /// Tests columns 1-100 to include multi-character column names
    /// </summary>
    [Benchmark]
    public void TransformToCharacterIndex_LargeIndices_WithInlining()
    {
        for (int row = 0; row < 100; row++)
        {
            for (uint col = 1; col <= 100; col++)
            {
                var cellRef = TransformToCharacterIndexInlined(col);
                _ = cellRef;
            }
        }
    }
}
