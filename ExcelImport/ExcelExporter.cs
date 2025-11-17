using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Attributes;
using ExcelImport.Exceptions;
using ExcelImport.Package;
using ExcelImport.Spreadsheet;
using ExcelImport.Utils;

namespace ExcelImport;

/// <summary>
/// Class to export items to Excel using OpenXML.
/// </summary>
public class ExcelExporter
{    
    /// <summary>
    /// Export the given items to the given worksheets
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items">Items to export</param>
    /// <param name="filePath">The path of the exported file</param>
    public void ExportSheetsWithWriter<T>(List<(string, List<T>)> items, string filePath) where T : class
    {
        using var stream = new FileStream(filePath, FileMode.Create);
        using var spreadsheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        ExcelPackageHelper.GenerateWorkbook(spreadsheet);
        uint index = 1;
        foreach (var (sheetName, sheetItems) in items)
        {
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.AddNewPart<WorksheetPart>();
            ExportCoreWithWriter(sheetItems, worksheetPart, CollectionsMarshal.AsSpan(PropertyCache.GetCachedProperties(typeof(T))));
            ExcelPackageHelper.AddWorksheet(spreadsheet, worksheetPart, sheetName, index++);
        }
    }

    /// <summary>
    /// Export the given items to the given file using <see cref="OpenXmlWriter"/>"/>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items"></param>
    /// <param name="fileName"></param>
    public void ExportWithWriter<T>(List<T> items, string fileName) where T : class
    {
        using var stream = new FileStream(fileName, FileMode.Create);
        ExportWithWriter(items, stream);
    }

    /// <summary>
    /// Export the given items to the given streamn using <see cref="OpenXmlWriter"/>"/>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items"></param>
    /// <param name="stream"></param>
    public void ExportWithWriter<T>(List<T> items, Stream stream) where T : class
    {
        using var spreadsheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        ExcelPackageHelper.GenerateWorkbook(spreadsheet);
        WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.AddNewPart<WorksheetPart>();
        ExportCoreWithWriter(items, worksheetPart, CollectionsMarshal.AsSpan(PropertyCache.GetCachedProperties(typeof(T))));
        ExcelPackageHelper.AddWorksheet(spreadsheet, worksheetPart, "Export", 1);
    }
        
    /// <summary>
    /// Export the given items to the given worksheet using <see cref="OpenXmlWriter"/>"/>
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items"></param>
    /// <param name="worksheetPart"></param>
    /// <param name="properties"></param>
    /// <exception cref="ImportException"></exception>
    private void ExportCoreWithWriter<T>(
    List<T> items,
    WorksheetPart worksheetPart,
    ReadOnlySpan<PropertyImportInfo> properties
    ) where T : class
    {
        using var writer = OpenXmlWriter.Create(worksheetPart);
        writer.WriteStartElement(new Worksheet());
        writer.WriteStartElement(new SheetDimension { Reference = $"A1:{SheetHelper.TransformToCharacterIndex((uint)properties.Length)}{items.Count + 1}" });
        writer.WriteEndElement(); // SheetDimension
        writer.WriteStartElement(new SheetData());

        // Header
        writer.WriteStartElement(new Row { RowIndex = 1 });
        for (int i = 0; i < properties.Length; i++)
        {
            var colRef = new SheetCell(SheetHelper.TransformToCharacterIndex((uint)i + 1), 1).ToString();
            writer.WriteStartElement(new Cell
            {
                CellReference = colRef,
                DataType = CellValues.String
            });
            writer.WriteElement(new CellValue(properties[i].ColumnIdentifier));
            writer.WriteEndElement(); // Cell
        }
        writer.WriteEndElement(); // Row

        // Data
        uint rowIndex = 2;
        foreach (var item in items)
        {
            writer.WriteStartElement(new Row { RowIndex = rowIndex });

            for (int i = 0; i < properties.Length; i++)
            {
                var prop = properties[i];
                var value = prop.Property.GetValue(item);

                if (prop.Required && value is null)
                    throw new ImportException(i.ToString(), prop.ColumnIdentifier);

                if (value is null)
                    continue;

                var cellValue = prop.TypeConverter.ConvertToCellValue(value);
                if (cellValue is null)
                    continue;

                var cellRef = new SheetCell(SheetHelper.TransformToCharacterIndex((uint)i + 1), (int)rowIndex).ToString();

                writer.WriteStartElement(new Cell
                {
                    CellReference = cellRef,
                    DataType = prop.TypeConverter.CellType,
                    StyleIndex = prop.TypeConverter.StyleIndex
                });
                writer.WriteElement(cellValue);
                writer.WriteEndElement(); // Cell
            }

            writer.WriteEndElement(); // Row
            rowIndex++;
        }

        writer.WriteEndElement(); // SheetData
        writer.WriteEndElement(); // Worksheet
        writer.Close();
    }
}