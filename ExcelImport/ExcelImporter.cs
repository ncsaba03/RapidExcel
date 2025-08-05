using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport;
using ExcelImport.Exceptions;
using ExcelImport.Spreadsheet;
using ExcelImport.Utils;

/// <summary>
/// Imports data from an Excel file 
/// </summary>
public class ExcelImporter
{

    /// <summary>
    /// Imports the data from the Excel file.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="filePath"></param>
    /// <param name="headerRowIndex"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    /// <exception cref="ImportException"></exception>
    public IEnumerable<T> Import<T>(string filePath, uint headerRowIndex = 0)
        where T : new()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(T));
        var context = new ExcelImportContext(filePath);

        return ImportCore<T>(context, properties.ToDictionary(t => t.ColumnIdentifier), headerRowIndex);
    }

    /// <summary>
    /// Imports the data from the worksheet.
    /// </summary>
    /// <typeparam name="T"></typeparam> 
    /// <param name="context"></param>
    /// <param name="attributes"></param>    
    /// <param name="headerRowIndex"></param>
    /// <returns></returns>
    /// <exception cref="ImportException"></exception>
    private IEnumerable<T> ImportCore<T>(ExcelImportContext context, IReadOnlyDictionary<string, PropertyImportInfo> attributes, uint headerRowIndex)
        where T : new()
    {
        var headerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        uint rowindex = 0;
        using var reader = OpenXmlReader.Create(context.WorksheetPart);
        while (reader.Read())
        {
            if (reader.ElementType == typeof(Row) && reader.IsStartElement)
            {
                if (reader.LoadCurrentElement() is not Row row || (row.RowIndex?.Value ?? headerRowIndex) < headerRowIndex)
                {
                    continue;
                }

                if (row.RowIndex?.Value == headerRowIndex)
                {
                    headerMap = GetHeaders(row, context.SharedStrings);
                    continue;
                }

                rowindex++;
                var item = new T();

                foreach (var cell in row.Elements<Cell>())
                {
                    var col = SheetHelper.GetColumnIndexFromCellReference(cell.CellReference!.Value).ToString();
                    var value = GetCellValue(cell, context.SharedStrings);

                    if (!headerMap.TryGetValue(col, out var headerName))
                    {
                        continue;
                    }

                    if (!attributes.TryGetValue(headerName, out var prop))
                    {
                        continue;
                    }

                    if (!headerName.Equals(prop.ColumnIdentifier, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    if (prop.Property.PropertyType == typeof(string))
                    {
                        if (prop.Required && value.AsSpan().IsEmpty)
                        {
                            var message = $"{prop.ColumnIdentifier} is required!";
                            throw new ImportException(cell.CellReference.Value ?? string.Empty, message);
                        }

                        prop.Property.SetValue(item, value);

                        continue;
                    }

                    object? convertedValue = null;

                    if (prop.TypeConverter is not null)
                    {
                        convertedValue = prop.TypeConverter.Convert(value);
                    }

                    if (prop.Required && convertedValue is null)
                    {
                        var message = $"{prop.ColumnIdentifier} is required!";
                        throw new ImportException(cell.CellReference.Value ?? string.Empty, message);
                    }

                    prop.Property.SetValue(item, convertedValue);

                }

                yield return item;
            }
        }

        context?.Dispose();
    }

    /// <summary>
    /// Gets the headers from the first row of the worksheet.
    /// </summary>
    /// <param name="headerRow"></param>
    /// <param name="sharedStrings"></param>
    /// <returns></returns>
    private Dictionary<string, string> GetHeaders(Row headerRow, SharedStringTable sharedStrings)
    {
        var headers = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        foreach (var cell in headerRow.Elements<Cell>())
        {
            var column = SheetHelper.GetColumnIndexFromCellReference(cell.CellReference?.Value ?? "").ToString();
            var value = GetCellValue(cell, sharedStrings);
            headers[column] = value;
        }

        return headers;
    }

    /// <summary>
    /// Gets the cell value from the cell.
    /// </summary>
    /// <param name="cell"></param>
    /// <param name="sharedStrings"></param>
    /// <returns></returns>
    private static string GetCellValue(Cell cell, SharedStringTable sharedStrings)
    {
        if (cell == null || cell.CellValue == null)
            return string.Empty;

        if (cell.DataType?.Value != CellValues.SharedString)
        {
            return cell.CellValue.Text;
        }

        if (int.TryParse(cell.CellValue.Text, out var index))
        {
            return sharedStrings.ChildElements[index].InnerText;
        }

        return cell.CellValue.Text;
    }
}
