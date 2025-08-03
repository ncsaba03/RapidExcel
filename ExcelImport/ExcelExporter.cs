using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Exceptions;
using ExcelImport.Package;
using ExcelImport.Spreadsheet;

namespace ExcelImport;

public class ExcelExporter
{
    public static ConcurrentDictionary<int, string> SharedStringTable { get; } = new ConcurrentDictionary<int, string>();


    /// <summary>
    /// Export the given items to the given worksheets
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items"></param>
    /// <param name="filePath"></param>
    public void ExportSheets<T>(List<(string, List<T>)> items, string filePath) where T : class
    {
        using var stream = new FileStream(filePath, FileMode.Create);
        using var spreadsheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        ExcelPackageHelper.GenerateWorkbook(spreadsheet);

        uint index = 1;
        foreach (var (sheetName, sheetItems) in items)
        {
            var sheetData = new SheetData();
            SheetRange sheetRange = ExportCore(sheetItems, sheetData);
            CreateWorksheet(spreadsheet, sheetData, sheetName, index++, sheetRange);
        }
    }

    /// <summary>
    /// Export the given items to the given worksheets
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items"></param>
    /// <param name="filePath"></param>
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
            AddWorksheet(spreadsheet, worksheetPart, sheetName, index++);
        }
    }

    /// <summary>
    /// Export the given items to the given file
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items"></param>
    /// <param name="filePath"></param>
    public void Export<T>(List<T> items, string filePath) where T : class
    {
        using var stream = new FileStream(filePath, FileMode.Create);
        Export(items, stream);
    }

    /// <summary>
    /// Export the given items to the given stream
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
        AddWorksheet(spreadsheet, worksheetPart, "Export", 1);

    }

    /// <summary>
    /// Export the given items to the given file with a writer
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
    /// Export the given items to the given stream
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items"></param>
    /// <param name="stream"></param>
    public void Export<T>(List<T> items, Stream stream) where T : class
    {
        using var spreadsheet = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
        ExcelPackageHelper.GenerateWorkbook(spreadsheet);

        var sheetData = new SheetData();
        SheetRange sheetRange = ExportCore(items, sheetData);

        CreateWorksheet(spreadsheet, sheetData, "Export", 1, sheetRange);
    }

    /// <summary>
    /// Export the given items to the given sheet
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="items"></param>
    /// <param name="sheetData"></param>
    /// <returns></returns>
    private SheetRange ExportCore<T>(List<T> items, SheetData sheetData) where T : class
    {
        var properties = CollectionsMarshal.AsSpan(PropertyCache.GetCachedProperties(typeof(T)));

        WriteHeader(sheetData, properties);
        WriteData(sheetData, items, properties);

        uint columns = (uint)properties.Length;
        var rows = items.Count + 1;
        var sheetRange = new SheetRange("A1", $"{SheetHelper.GetColumnName(columns)}{rows}");
        return sheetRange;
    }

    /// <summary>
    /// Write the header of the sheet
    /// </summary>
    /// <param name="sheetData"></param>
    /// <param name="properties"></param>
    private void WriteHeader(SheetData sheetData, ReadOnlySpan<PropertyImportInfo> properties)
    {
        var row = new Row { RowIndex = 1 };

        for (int i = 0; i < properties.Length; i++)
        {
            var prop = properties[i];
            var cell = new Cell
            {
                CellReference = new SheetCell(SheetHelper.GetColumnName((uint)i + 1), 1).ToString(),
                DataType = CellValues.String,
                CellValue = new CellValue(prop.ColumnIdentifier)
            };
            row.AppendChild(cell);
        }

        sheetData.AppendChild(row);
    }

    /// <summary>
    /// Write the data of the sheet
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <param name="sheetData"></param>
    /// <param name="items"></param>
    /// <param name="properties"></param>
    /// <exception cref="ImportException"></exception>
    private void WriteData<T>(SheetData sheetData, List<T> items, ReadOnlySpan<PropertyImportInfo> properties)
    {
        uint rowIndex = 2;
        foreach (var item in items)
        {
            var row = new Row { RowIndex = rowIndex };

            for (int i = 0; i < properties.Length; i++)
            {
                var prop = properties[i];
                var value = prop.Property.GetValue(item);

                if (prop.Required && value is null)
                {
                    throw new ImportException(i.ToString(), prop.ColumnIdentifier);
                }

                if (value is null)
                {
                    continue;
                }

                var cellValue = prop.TypeConverter.ConvertToCellValue(value);

                if (cellValue is null)
                {
                    continue;
                }

                var sheetCell = new SheetCell(SheetHelper.GetColumnName((uint)i + 1), (int)rowIndex);

                var cell = new Cell
                {
                    CellReference = sheetCell.ToString(),
                    DataType = prop.TypeConverter.CellType,
                    CellValue = cellValue,
                    StyleIndex = prop.TypeConverter.StyleIndex,
                };

                row.AppendChild(cell);
            }

            sheetData.AppendChild(row);
            rowIndex++;
        }

    }

    /// <summary>
    /// Create a new worksheet in the given spreadsheet
    /// </summary>
    /// <param name="spreadsheet"></param>
    /// <param name="sheetData"></param>
    /// <param name="sheetName"></param>
    /// <param name="sheetId"></param>
    /// <param name="sheetRange"></param>
    private static void AddWorksheet(SpreadsheetDocument spreadsheet, WorksheetPart worksheetPart, string sheetName, uint sheetId)
    {
        var workbookPart = spreadsheet.WorkbookPart!;

        var sheets = spreadsheet.WorkbookPart!.Workbook.Sheets;

        if (sheets == null)
        {
            sheets = new Sheets();
            spreadsheet.WorkbookPart.Workbook.AppendChild(sheets);
        }

        sheets.Append(new Sheet
        {
            Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = sheetName
        });
    }

    /// <summary>
    /// Create a new worksheet in the given spreadsheet
    /// </summary>
    /// <param name="spreadsheet"></param>
    /// <param name="sheetData"></param>
    /// <param name="sheetName"></param>
    /// <param name="sheetId"></param>
    /// <param name="sheetRange"></param>
    private static void CreateWorksheet(SpreadsheetDocument spreadsheet, SheetData sheetData, string sheetName, uint sheetId, SheetRange sheetRange)
    {
        var workbookPart = spreadsheet.WorkbookPart!;

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(sheetData);

        var sheets = spreadsheet.WorkbookPart!.Workbook.Sheets;

        if (sheets == null)
        {
            sheets = new Sheets();
            spreadsheet.WorkbookPart.Workbook.AppendChild(sheets);
        }

        sheets.Append(new Sheet
        {
            Id = spreadsheet.WorkbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = sheetName
        });

        var sheetDimension = new SheetDimension
        {
            Reference = new StringValue(sheetRange.ToString())
        };
        worksheetPart.Worksheet.SheetDimension = sheetDimension;
    }

    /// <summary>
    /// Export the given items to the given worksheet
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
        writer.WriteStartElement(new SheetDimension { Reference = $"A1:{SheetHelper.GetColumnName((uint)properties.Length)}{items.Count + 1}" });
        writer.WriteEndElement(); // SheetDimension
        writer.WriteStartElement(new SheetData());

        // Header
        writer.WriteStartElement(new Row { RowIndex = 1 });
        for (int i = 0; i < properties.Length; i++)
        {
            var colRef = new SheetCell(SheetHelper.GetColumnName((uint)i + 1), 1).ToString();
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

                var cellRef = new SheetCell(SheetHelper.GetColumnName((uint)i + 1), (int)rowIndex).ToString();

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