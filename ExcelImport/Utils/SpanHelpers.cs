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

}
