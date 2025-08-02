namespace ExcelImport.Spreadsheet
{
    public static class SheetHelper
    {
        /// <summary>
        /// Gets the column index from a string
        /// </summary>
        /// <param name="col"></param>
        /// <returns></returns>
        public static int GetColumnIndex(ReadOnlySpan<char> col)
        {            
            int index = 0;
            for (int i = 0; i < col.Length; i++)
            {
                char ch = col[i];
                // transform character into base 26 number
                int value = ch - 'A' + 1;
                index = index * 26 + value;
            }
            return index;
        }
        public static string GetColumnName(uint colIndex)
        {
            if (colIndex == 0)
                throw new ArgumentOutOfRangeException(nameof(colIndex), "Column index must be >= 1");

            Span<char> buffer = stackalloc char[8]; 
            int pos = buffer.Length;

            uint index = colIndex;
            while (index > 0)
            {
                index--; 
                buffer[--pos] = (char)('A' + (index % 26));
                index /= 26;
            }

            return new(buffer[pos..]);
        }
    }
}
