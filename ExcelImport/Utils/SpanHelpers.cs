using System.Globalization;

namespace ExcelImport.Utils;

public static class SpanHelpers
{
    /// <summary>
    /// Parses a decimal number from a <see cref="ReadOnlySpan{char}"/>. This method is optimized to avoid heap allocations.
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
