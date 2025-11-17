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

    public override int GetHashCode()
    {
        return HashCode.Combine(Col, Row);
    }

    public bool Equals(SheetCell other)
    {
        if (Col == null || other.Col == null)
            return false;
        
        return Col == other.Col && Row == other.Row;
    }

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
