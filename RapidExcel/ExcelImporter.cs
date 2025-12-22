using System.Xml;
using RapidExcel.Attributes;
using RapidExcel.Exceptions;
using RapidExcel.Spreadsheet;
using RapidExcel.Utils;

namespace RapidExcel;

/// <summary>
/// Imports data from an Excel file 
/// </summary>
public class ExcelImporter
{

    /// <summary>
    /// Imports the data from the Excel file.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="filePath">The filepath</param>
    /// <param name="headerRowIndex">The index of the header row</param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    /// <exception cref="ImportException"></exception>
    public IEnumerable<T> Import<T>(string filePath, uint headerRowIndex = 1)
        where T : new()
    {
        var properties = PropertyCache.GetCachedProperties(typeof(T));
        using var context = new ExcelImportContext(filePath);
        if (headerRowIndex < 1)
        {
            throw new InvalidOperationException("Header row index must be greater than 0.");
        }

        var attributes = properties.ToDictionary(p => p.ColumnIdentifier, StringComparer.OrdinalIgnoreCase);

        foreach (var item in ImportCoreSax<T>(context, attributes, headerRowIndex))
        {
            yield return item;
        }
    }

    /// <summary>
    /// SAX-based import implementation
    /// </summary>
    private IEnumerable<T> ImportCoreSax<T>(ExcelImportContext context, IReadOnlyDictionary<string, PropertyImportInfo> attributes, uint headerRowIndex)
         where T : new()
    {
        var headerMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        using var reader = XmlReader.Create(context.WorksheetPart.GetStream());

        while (reader.Read())
        {
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "row")
            {
                continue;
            }           

            uint? rowIndex = uint.TryParse(reader.GetAttribute("r"), out var currentRowIndex) ? currentRowIndex : null;
            if (!rowIndex.HasValue)
            {
                continue;
            }

            if (currentRowIndex < headerRowIndex)
            {
                continue;
            }

            T? currentItem = default;

            if (currentRowIndex > headerRowIndex)
            {
                currentItem = new T();
            }

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "row")
                {
                    break;
                }

                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                {
                    var cellReference = reader.GetAttribute("r") ?? string.Empty;
                    var value = GetCellValue(reader, context);
                    var colIndex = SheetHelper.GetColumnIndexFromCellReference(cellReference).ToString();

                    if (currentRowIndex == headerRowIndex)
                    {
                        if (string.IsNullOrWhiteSpace(value))
                        {
                            continue;
                        }

                        headerMap[colIndex] = value;
                        continue;
                    }

                    if (currentItem != null && headerMap.TryGetValue(colIndex, out var headerName))
                    {
                        SetProperty(currentItem, headerName, value, cellReference, attributes);
                    }
                }
            }

            if (currentItem != null)
            {
                yield return currentItem;
                currentItem = default;
            }
        }
    }

    /// <summary>
    /// Gets the cell value from the XML reader
    /// </summary>
    /// <param name="reader">XmlReader</param>
    /// <param name="context">ExcelImportContext</param>
    /// <returns></returns>
    private static string GetCellValue(XmlReader reader, ExcelImportContext context)
    {        
        bool isSharedString = reader.GetAttribute("t") == "s";
        string rawValue = string.Empty;

        if (reader.IsEmptyElement)
        {
            return rawValue;
        }

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "c")
            {
                break;
            }

            if (reader.NodeType == XmlNodeType.Element &&
               (reader.LocalName == "v" || reader.LocalName == "t"))
            {
                if (reader.IsEmptyElement) continue;

                if (reader.Read())
                {
                    if (reader.NodeType is XmlNodeType.Text or
                                           XmlNodeType.CDATA or
                                           XmlNodeType.Whitespace)
                    {
                        rawValue = reader.Value;
                    }
                }
            }
        }

        return isSharedString && int.TryParse(rawValue, out int id) ? context.GetSharedString(id) : rawValue;

    }

    /// <summary>
    /// Sets the property value on the item
    /// </summary>
    /// <typeparam name="T">Type</typeparam>
    /// <param name="item">The item</param>
    /// <param name="headerName">The name of header</param>
    /// <param name="value">The value</param>
    /// <param name="cellRef">The reference of the cell</param>
    /// <param name="attributes">The attribute cache</param>
    /// <exception cref="ImportException"></exception>
    private static void SetProperty<T>(T item, string headerName, string value, string cellRef, IReadOnlyDictionary<string, PropertyImportInfo> attributes)
    {
        if (!attributes.TryGetValue(headerName, out var prop)) return;
        if (!headerName.Equals(prop.ColumnIdentifier, StringComparison.OrdinalIgnoreCase)) return;

        if (prop.Property.PropertyType == typeof(string))
        {
            if (prop.Required && string.IsNullOrWhiteSpace(value))
            {
                throw new ImportException(cellRef, $"{prop.ColumnIdentifier} is required!");
            }
            prop.Property.SetValue(item, value);
            return;
        }

        object? convertedValue = null;
        if (prop.TypeConverter is not null)
        {
            convertedValue = prop.TypeConverter.Convert(value);
        }

        if (prop.Required && convertedValue is null)
        {
            throw new ImportException(cellRef, $"{prop.ColumnIdentifier} is required!");
        }

        prop.Property.SetValue(item, convertedValue);
    }
}

