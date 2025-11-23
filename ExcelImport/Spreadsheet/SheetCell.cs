using System.Runtime.InteropServices;

namespace ExcelImport.Spreadsheet;

/// <summary>
///  Represents a Google sheet cell
/// </summary>
[StructLayout(LayoutKind.Auto)]
public  struct SheetCell : IEquatable<SheetCell>, IComparable<SheetCell>
{
    /// <summary>
    /// Represents the columns identifier
    /// </summary>
    public readonly int ColIndex => SheetHelper.GetColumnIndex(Col);
    
    /// <summary>
    /// Gets the columns identifier
    /// </summary>
    public string Col { get; }

    /// <summary>
    /// Gets the rows identifier
    /// </summary>
    public int Row { get; }

    /// <summary>
    /// Parse a <c>SheetCell</c> object from string
    /// </summary>
    /// <param name="cellString"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentException"></exception>
    public static SheetCell Parse(ReadOnlySpan<char> cellString)
    {
        var cell = cellString.Trim();
        int colIndex = -1;
        for (int i = 0; i < cell.Length; i++)
        {
            if (char.IsDigit(cell[i]))
            {
                colIndex = i;
                break;
            }
        }

        if (colIndex == -1) throw new ArgumentException(nameof(cell), "There is no row indicator in the string");

        var col = cell[..colIndex];
        int row = int.Parse(cell[colIndex..]);

        return new(col.ToString(), row);
    }

    /// <summary>
    /// Add the given number of rows to the cell
    /// </summary>
    /// <param name="count"></param>
    /// <returns cref="SheetCell"></returns>
    /// <exception cref="ArgumentException"></exception>
    public SheetCell AddRows(int count)
    {
        if (Row + count < 0) throw new ArgumentException(nameof(count));
        return new(Col, Row + count);
    }

    /// <summary>
    /// Adds the given number of columns to the rows
    /// </summary>
    /// <param name="count"></param>
    /// <returns cref="SheetCell"></returns>
    public SheetCell AddColumns(int count)
    {
        uint currentIndex = (uint)SheetHelper.GetColumnIndex(Col);
        uint newIndex = currentIndex + (uint)count;
        string newCol = SheetHelper.TransformToCharacterIndex(newIndex);
        return new SheetCell(newCol, Row);
    }

    /// <summary>
    /// Construct <see cref="SheetCell"/>
    /// </summary>
    /// <param name="col"></param>
    /// <param name="row"></param>
    /// <exception cref="ArgumentNullException"></exception>
    public SheetCell(string col, int row)
    {
        Row = row;
        Col = col ?? throw new ArgumentNullException(nameof(col));
    }

    /// <summary>
    /// Returns the string representation of the sheet cell 
    /// </summary>
    /// <returns></returns>
    public override string ToString()
    {
        return $"{Col}{Row}";
    }

    #region CompareTo and Equals

    /// <summary>
    /// Serves as the default hash function for the object.
    /// </summary>
    /// <remarks>The hash code is computed based on the values of the Col and Row properties. This ensures
    /// that objects with the same Col and Row values produce the same hash code, which is important for correct
    /// behavior in hash-based collections such as dictionaries and hash sets.</remarks>
    /// <returns>A 32-bit signed integer hash code that represents the current object.</returns>
    public override int GetHashCode()
    {
        return HashCode.Combine(Col, Row);
    }

    /// <summary>
    /// Determines whether the current SheetCell is equal to the specified SheetCell based on column and row values.
    /// </summary>
    /// <param name="other">The SheetCell to compare with the current SheetCell. The comparison is based on the Col and Row properties.</param>
    /// <returns>true if both SheetCell instances have non-null Col values and their Col and Row properties are equal; otherwise,
    /// false.</returns>
    public bool Equals(SheetCell other)
    {
        if (Col == null || other.Col == null)
            return false;
        
        return Col == other.Col && Row == other.Row;
    }

    /// <summary>
    /// Compares the current SheetCell instance to another SheetCell and returns an integer that indicates their
    /// relative position in a sort order.
    /// </summary>
    /// <remarks>Comparison is performed using ordinal string comparison for the Col property. If either Col
    /// property is null, the current instance is considered less than <paramref name="other"/>.</remarks>
    /// <param name="other">The SheetCell to compare with the current instance. The comparison is based first on the Col property, then on
    /// the Row property if the columns are equal.</param>
    /// <returns>A value less than zero if the current instance precedes <paramref name="other"/> in the sort order; zero if they
    /// are equal; or a value greater than zero if the current instance follows <paramref name="other"/>.</returns>
    public int CompareTo(SheetCell other)
    {
        if (other.Col == null || Col == null)
            return -1;
        var colDiff = string.Compare(Col, other.Col, StringComparison.Ordinal);
        if (colDiff != 0)
            return colDiff;

        return Row.CompareTo(other.Row);
    }

    #endregion
}
