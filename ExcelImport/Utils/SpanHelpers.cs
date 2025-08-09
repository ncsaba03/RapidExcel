using System.Globalization;
using System.Runtime.CompilerServices;

namespace ExcelImport.Utils;

/// <summary>
/// Utility class for parsing decimal numbers from spans.
/// </summary>
public static class SpanHelpers
{
     /// <summary>
    /// Interns a string if it is safe to do so. A string is considered safe if it is not empty, does not exceed 256 characters, and is less than or equal to 64 characters in length.
    /// </summary>
    /// <param name="span"></param>
    /// <returns></returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    public static string? InternIfSafe(ReadOnlySpan<char> span)
    {
        if (span.IsEmpty)
        {
            return null;
        }

        if (span.Length <= 64)  
        {
            Span<char> buffer = stackalloc char[span.Length];
            span.CopyTo(buffer);
            return string.Intern(new string(buffer));
        }

        if (span.Length > 256)  
        {
            return null;
        }

        return string.Intern(span.ToString());
    }

}
