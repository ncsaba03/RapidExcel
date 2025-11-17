using ExcelImport.Utils;

namespace ExcelImport.Test;

public class UtilityTests
{
    #region SpanHelpers Tests

    [Fact]
    public void SpanHelpers_InternIfSafe_EmptySpan_ReturnsNull()
    {
        var result = SpanHelpers.InternIfSafe(ReadOnlySpan<char>.Empty);

        Assert.Null(result);
    }

    [Theory]
    [InlineData("a")]
    [InlineData("test")]
    [InlineData("Hello World")]
    public void SpanHelpers_InternIfSafe_ShortString_ReturnsInternedString(string input)
    {
        var result = SpanHelpers.InternIfSafe(input.AsSpan());

        Assert.NotNull(result);
        Assert.Equal(input, result);
        Assert.Same(string.Intern(input), result); // Verify it's actually interned
    }

    [Fact]
    public void SpanHelpers_InternIfSafe_256CharString_ReturnsInternedString()
    {
        var input = new string('a', 256);

        var result = SpanHelpers.InternIfSafe(input.AsSpan());

        Assert.NotNull(result);
        Assert.Equal(input, result);
    }

    [Fact]
    public void SpanHelpers_InternIfSafe_257CharString_ReturnsNull()
    {
        var input = new string('a', 257);

        var result = SpanHelpers.InternIfSafe(input.AsSpan());

        Assert.Null(result);
    }

    [Fact]
    public void SpanHelpers_InternIfSafe_LongString_ReturnsNull()
    {
        var input = new string('a', 500);

        var result = SpanHelpers.InternIfSafe(input.AsSpan());

        Assert.Null(result);
    }

    [Fact]
    public void SpanHelpers_InternIfSafe_SameString_ReturnsSameInstance()
    {
        var input = "test";

        var result1 = SpanHelpers.InternIfSafe(input.AsSpan());
        var result2 = SpanHelpers.InternIfSafe(input.AsSpan());

        Assert.Same(result1, result2);
    }

    [Theory]
    [InlineData("USD")]
    [InlineData("EUR")]
    [InlineData("HUF")]
    public void SpanHelpers_InternIfSafe_CommonValues_Interns(string currency)
    {
        var result = SpanHelpers.InternIfSafe(currency.AsSpan());

        Assert.NotNull(result);
        Assert.Same(string.Intern(currency), result);
    }

    #endregion

    #region SpanSplitEnumerator Tests

    [Fact]
    public void SpanSplitEnumerator_SingleSegment_NoSeparator_ReturnsWhole()
    {
        var input = "hello".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("hello".AsSpan()));
        Assert.False(enumerator.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_MultipleSegments_SplitsCorrectly()
    {
        var input = "a,b,c".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("a".AsSpan()));

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("b".AsSpan()));

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("c".AsSpan()));

        Assert.False(enumerator.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_EmptySpan_ReturnsNoElements()
    {
        var input = ReadOnlySpan<char>.Empty;
        var enumerator = new SpanSplitEnumerator(input, ',');

        Assert.False(enumerator.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_TrailingSeparator_ReturnsEmptyLast()
    {
        var input = "a,b,".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("a".AsSpan()));

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("b".AsSpan()));

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.IsEmpty);

        Assert.False(enumerator.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_LeadingSeparator_ReturnsEmptyFirst()
    {
        var input = ",a,b".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.IsEmpty);

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("a".AsSpan()));

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("b".AsSpan()));

        Assert.False(enumerator.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_ConsecutiveSeparators_ReturnsEmptySegments()
    {
        var input = "a,,b".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("a".AsSpan()));

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.IsEmpty);

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("b".AsSpan()));

        Assert.False(enumerator.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_OnlySeparators_ReturnsEmptySegments()
    {
        var input = ",,,".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        for (int i = 0; i < 4; i++)
        {
            Assert.True(enumerator.MoveNext());
            Assert.True(enumerator.Current.IsEmpty);
        }

        Assert.False(enumerator.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_Position_TracksCorrectly()
    {
        var input = "a,b,c".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        Assert.Equal(0, enumerator.Position);

        enumerator.MoveNext();
        Assert.Equal(2, enumerator.Position); // After 'a,'

        enumerator.MoveNext();
        Assert.Equal(4, enumerator.Position); // After 'a,b,'

        enumerator.MoveNext();
        Assert.Equal(6, enumerator.Position); // After 'a,b,c'
    }

    [Fact]
    public void SpanSplitEnumerator_ForeachLoop_Works()
    {
        var input = "one,two,three".AsSpan();
        var segments = new List<string>();

        foreach (var segment in new SpanSplitEnumerator(input, ','))
        {
            segments.Add(segment.ToString());
        }

        Assert.Equal(3, segments.Count);
        Assert.Equal("one", segments[0]);
        Assert.Equal("two", segments[1]);
        Assert.Equal("three", segments[2]);
    }

    [Fact]
    public void SpanSplitEnumerator_DifferentSeparators_WorksCorrectly()
    {
        var input = "apple;banana;cherry".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ';');

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("apple".AsSpan()));

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("banana".AsSpan()));

        Assert.True(enumerator.MoveNext());
        Assert.True(enumerator.Current.SequenceEqual("cherry".AsSpan()));

        Assert.False(enumerator.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_GetEnumerator_ReturnsSelf()
    {
        var input = "test".AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        var result = enumerator.GetEnumerator();

        // Should be the same struct (ref struct can't be compared with Assert.Same)
        Assert.True(enumerator.MoveNext());
        Assert.True(result.MoveNext());
    }

    [Fact]
    public void SpanSplitEnumerator_LargeInput_HandlesEfficiently()
    {
        var parts = new string[1000];
        for (int i = 0; i < 1000; i++)
            parts[i] = $"part{i}";

        var input = string.Join(",", parts).AsSpan();
        var enumerator = new SpanSplitEnumerator(input, ',');

        int count = 0;
        while (enumerator.MoveNext())
        {
            count++;
        }

        Assert.Equal(1000, count);
    }

    #endregion
}
