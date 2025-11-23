using System.Globalization;
using System.Runtime.CompilerServices;

namespace ExcelImport.Utils;

/// <summary>
/// Utility class for parsing decimal numbers from spans.
/// </summary>
public static class SpanHelpers
{
     /// <summary>
    /// Interns a string if it is safe to do so. A string is considered safe if it is not empty and does not exceed 256 characters.
    /// Returns null for empty spans or strings longer than 256 characters.
    /// </summary>
    /// <param name="span"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string? InternIfSafe(ReadOnlySpan<char> span)
    {
        // Only intern reasonably sized strings (empty or > 256 chars are skipped)
        if (span.IsEmpty || span.Length > 256)
        {
            return null;
        }

        return string.Intern(new string(span));
    }

    /// <summary>
    /// Parses a decimal number from a <see cref="System.ReadOnlySpan{T}"/> of char. This method is optimized to avoid heap allocations.
    /// </summary>
    /// <param name="input"></param>
    /// <param name="result"></param>
    /// <returns></returns>
    public static bool TryParseDecimal(ReadOnlySpan<char> input, out decimal result)
    {
        result = default;

        // Check if the input is empty or too long
        if (input.Length == 0 || input.Length > 36)
            return false;

        // Check if the input is a valid decimal number
        if (decimal.TryParse(input, CultureInfo.InvariantCulture, out result))
            return true;

        // try to remove the thousand separator without allocating on the heap
        Span<char> buffer = stackalloc char[input.Length];
        int bufferIndex = 0;
        bool skipped = false;

        for (int i = 0; i < input.Length; i++)
        {
            if (!skipped && (input[i] == ',' || input[i] == '.'))
            {
                skipped = true;
                continue;
            }

            buffer[bufferIndex++] = input[i];
        }

        return decimal.TryParse(buffer[..bufferIndex], CultureInfo.InvariantCulture, out result);
    }
}
