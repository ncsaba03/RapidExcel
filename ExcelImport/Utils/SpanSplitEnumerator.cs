namespace ExcelImport.Utils;

/// <summary>
/// Enumerator for splitting a ReadOnlySpan<char> by a specified separator.
/// </summary>
public ref struct SpanSplitEnumerator
{
    private readonly ReadOnlySpan<char> _span;
    private readonly char _separator;
    private int _pos;

    public SpanSplitEnumerator(ReadOnlySpan<char> span, char separator)
    {
        _span = span;
        _separator = separator;
        _pos = 0;
        Current = default;
    }

    public SpanSplitEnumerator GetEnumerator()
    {
        return this;
    }

    public readonly int Position => _pos;

    public ReadOnlySpan<char> Current { get; private set; }

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