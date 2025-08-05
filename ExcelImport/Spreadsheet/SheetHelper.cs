using System.Runtime.CompilerServices;

namespace ExcelImport.Spreadsheet;

/// <summary>
/// Helper class for working with spreadsheet columns.
/// </summary>
public static class SheetHelper
{
    /// <summary>
    /// Gets the column index from a string
    /// </summary>
    /// <param name="col"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static int GetColumnIndex(ReadOnlySpan<char> col)
    {            
        int index = 0;
        for (int i = 0; i < col.Length; i++)
        {
            char ch = col[i];
            // transform character into base 26 number
            int value = ch - 'A' + 1;
            index = index * 26 + value;
        }
        return index;
    }

    /// <summary>
    /// Transforms a column index to the corresponding Excel column name. <para>(e.g., 1 -> "A", 2 -> "B", ..., 27 -> "AA")</para>
    /// </summary>
    /// <param name="colIndex"></param>
    /// <returns>
    /// The excel column name corresponding to the index.
    /// </returns>
    /// <exception cref="ArgumentOutOfRangeException"></exception>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string TransformToCharacterIndex(uint colIndex)
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
    /// Extracts the column index from a cell reference.
    /// <para>(e.g., "A1" -> "A", "B2" -> "B")</para>
    /// </summary>
    /// <param name="cellRef"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static ReadOnlySpan<char> GetColumnIndexFromCellReference(ReadOnlySpan<char> cellRef)
    {
        for (int i = 0; i < cellRef.Length; i++)
        {
            if (char.IsDigit(cellRef[i]))
            {
                return cellRef[..i];
            }
        }

        return cellRef;
    }
}
