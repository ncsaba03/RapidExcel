using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

namespace ExcelImport.Spreadsheet;

/// <summary>
/// Represents a SheetRange
/// </summary>
[StructLayout(LayoutKind.Auto)]
public struct SheetRange : IEquatable<SheetRange>, IComparable<SheetRange>, IEquatable<string>, IComparable<string>
{

    /// <summary>
    /// Construct a <see cref="SheetRange"/>
    /// </summary>
    /// <param name="from"></param>
    /// <param name="to"></param>
    public SheetRange(string from, string to)
    {

        From = SheetCell.Parse(from);
        To = SheetCell.Parse(to);
    }

    private int _columnLength = -1;

    /// <summary>
    /// Gets the length of the Columns in the range
    /// </summary>
    public int CoulumnLength
    {
        get
        {
            if (_columnLength != -1) return _columnLength;
            return To.ColIndex - From.ColIndex + 1;
        }
    }

    /// <summary>
    /// Construct a <see cref="SheetRange"/>
    /// </summary>
    /// <param name="from"></param>
    /// <param name="to"></param>
    public SheetRange(ReadOnlySpan<char> from, ReadOnlySpan<char> to)
    {
        From = SheetCell.Parse(from);
        To = SheetCell.Parse(to);
    }

    /// <summary>
    /// Checks that the given string is valid for range 
    /// </summary>
    /// <param name="range"></param>
    /// <returns></returns>
    public static bool IsValidForRange(ReadOnlySpan<char> range)
    {
        if (range.IsEmpty)
            return false;
        return Regex.IsMatch(range, "[A-Z]+[0-9]+[:][A-Z]+[0-9]+");
    }

    /// <summary>
    /// Add rows to the from part of the range 
    /// </summary>
    /// <param name="count"></param>
    /// <returns cref="SheetRange"></returns>
    public SheetRange AddRowsToBegining(int count)
    {
        var @new = new SheetRange(From.AddRows(count).ToString(), To.AddRows(count).ToString());
        Validate(@new);
        return @new;
    }

    /// <summary>
    /// Add rows to the end part of the range 
    /// </summary>
    /// <param name="count"></param>
    /// <returns cref="SheetRange"></returns>
    public SheetRange AddRowsToEnd(int count) => SetTo(To.AddRows(count).ToString());

    /// <summary>
    /// Add columns to the from part of the range 
    /// </summary>
    /// <param name="count"></param>
    /// <returns cref="SheetRange"></returns>
    public SheetRange AddColumnsToBegining(int count)
    {
        var @new = new SheetRange(From.AddColumns(count).ToString(), To.AddColumns(count).ToString());
        Validate(@new);
        return @new;
    }

    /// <summary>
    /// Add columns to the end part of the range 
    /// </summary>
    /// <param name="count"></param>
    /// <returns cref="SheetRange"></returns>
    public SheetRange AddColumnsToEnd(int count) => SetTo(To.AddColumns(count).ToString());

    /// <summary>
    /// Parse the <see cref="SheetRange"/> from the given string
    /// </summary>
    /// <param name="range"></param>
    /// <returns cref="SheetRange"></returns>
    /// <exception cref="ArgumentException"></exception>
    public static SheetRange Parse(ReadOnlySpan<char> range)
    {
        if (!IsValidForRange(range)) throw new ArgumentException(nameof(range), "Not a range!");
        var indexOf = range.IndexOf(':');
        var from = range[..indexOf];
        var to = range[(indexOf + 1)..];

        return new SheetRange(from, to);
    }

    /// <summary>
    /// Checks the given range is valid
    /// </summary>
    public bool IsValid => IsValidForRange(Range);

    private static void Validate(SheetRange range)
    {
        if (!range.IsValid) throw new ArgumentException(nameof(range));
    }

    /// <summary>
    /// Sets the range from part
    /// </summary>
    /// <param name="from"></param>
    /// <returns cref="SheetRange"></returns>
    public SheetRange SetFrom(string from)
    {
        var @new = new SheetRange(from, To.ToString());
        Validate(@new);
        return @new;
    }

    /// <summary>
    /// Sets the range end part
    /// </summary>
    /// <param name="to"></param>
    /// <returns cref="SheetRange"></returns>
    public SheetRange SetTo(string to)
    {
        var @new = new SheetRange(From.ToString(), to);
        Validate(@new);
        return @new;
    }

    /// <summary>
    /// Gets the range from part
    /// </summary>
    public SheetCell From { get; }

    /// <summary>
    /// Gets the range end part
    /// </summary>
    public SheetCell To { get; }

    /// <summary>
    /// Gets the string representation of the <see cref="SheetRange"/>
    /// </summary>
    private string Range => $"{From}:{To}";

    private int _rowLength = -1;

    /// <summary>
    /// Gets the length of the Rows in the range
    /// </summary>
    public int RowLength
    {
        get
        {
            if (_rowLength != -1) return _rowLength;
            return To.Row - From.Row + 1;
        }
    }

    /// <summary>
    /// Gets the string representation of the <see cref="SheetRange"/>
    /// </summary>
    /// <returns></returns>
    public override string ToString()
    {
        return Range;
    }

    /// <summary>
    /// Checks if the given cell is in the range
    /// </summary>
    /// <param name="cell"></param>
    /// <returns></returns>
    public bool IsInRange(SheetCell cell)
    {
        return cell.ColIndex >= From.ColIndex && cell.ColIndex <= To.ColIndex &&
               cell.Row >= From.Row && cell.Row <= To.Row;
    }

    /// <summary>
    /// Checks if the given cell is in the range
    /// </summary>
    /// <param name="cell"></param>
    /// <returns></returns>
    public bool IsInRange(ReadOnlySpan<char> cellString)
    {
        return IsInRange(SheetCell.Parse(cellString));
    }

    #region CompareTo and Equals

    /// <summary>
    /// Gets the hash code of the <see cref="SheetRange"/>
    /// </summary>
    /// <returns></returns>
    public override int GetHashCode()
    {
        return HashCode.Combine(From, To);
    }

    /// <summary>
    /// Checks if the <see cref="SheetRange"/> is equal to the given string
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(string? other)
    {
        return other != null && Range.Equals(other, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Compares the <see cref="SheetRange"/> with the given string
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public int CompareTo(string? other)
    {
        if (other == null)
            return 1;

        var range = Parse(other);
        return CompareTo(range);
    }

    /// <summary>
    /// Compares the two ranges
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public int CompareTo(SheetRange other)
    {
        var from = From.CompareTo(other.From);
        if (from != 0)
            return from;

        return To.CompareTo(other.To);
    }

    /// <summary>
    /// Checks if the two ranges are equal
    /// </summary>
    /// <param name="other"></param>
    /// <returns></returns>
    public bool Equals(SheetRange other)
    {
        return From.Equals(other.From) && To.Equals(other.To);
    }

    #endregion
}
