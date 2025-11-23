namespace ExcelImport.Utils;

/// <summary>
/// Enumerator for splitting a ReadOnlySpan&lt;char&gt; by a specified separator.
/// </summary>
public ref struct SpanSplitEnumerator
{
    private readonly ReadOnlySpan<char> _span;
    private readonly char _separator;
    private int _pos;

    /// <summary>
    /// Initializes a new instance of the SpanSplitEnumerator to iterate over substrings in the specified span, split by
    /// the given separator character.
    /// </summary>
    /// <param name="span">The read-only span of characters to be split into substrings.</param>
    /// <param name="separator">The character used to separate substrings within the span.</param>
    public SpanSplitEnumerator(ReadOnlySpan<char> span, char separator)
    {
        _span = span;
        _separator = separator;
        _pos = 0;
        Current = default;
    }

    /// <summary>
    /// Returns an enumerator that iterates through the spans produced by the splitter.
    /// </summary>
    /// <returns>A <see cref="SpanSplitEnumerator"/> that can be used to iterate through the split spans.</returns>
    public SpanSplitEnumerator GetEnumerator()
    {
        return this;
    }

    /// <summary>
    /// Gets the current position within the underlying data source.
    /// </summary>
    public readonly int Position => _pos;

    /// <summary>
    /// Gets the current read-only span of characters being processed or examined.
    /// </summary>
    public ReadOnlySpan<char> Current { get; private set; }

    /// <summary>
    /// Advances the enumerator to the next segment in the sequence, if one is available.
    /// </summary>
    /// <remarks>After calling MoveNext, the Current property contains the next segment in the sequence. If
    /// the end of the sequence has been reached, MoveNext returns false and Current is undefined.</remarks>
    /// <returns>true if the enumerator was successfully advanced to the next segment; otherwise, false.</returns>
    public bool MoveNext()
    {
        if (_span.IsEmpty || (uint)_pos > (uint)_span.Length) return false;

        int index = _span[_pos..].IndexOf(_separator);
        if (index == -1)
        {
            Current = _span[_pos..];
            _pos = _span.Length + 1;
            return true;
        }

        Current = _span.Slice(_pos, index);
        _pos += index + 1;
        return true;
    }
}