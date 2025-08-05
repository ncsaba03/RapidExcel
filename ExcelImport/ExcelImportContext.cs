using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Spreadsheet;

namespace ExcelImport;

/// <summary>
/// ExcelImportContext provides a context for importing data from an Excel file.
/// </summary>
public sealed class ExcelImportContext : IDisposable
{
    private readonly SpreadsheetDocument _document;
    private bool _disposed = false;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelImportContext"/> class with the specified file path.
    /// </summary>
    /// <param name="path"></param>
    /// <exception cref="InvalidOperationException"></exception>
    public ExcelImportContext(string path)
    {
        _document = SpreadsheetDocument.Open(path, false);
        WorkbookPart = _document.WorkbookPart ?? 
            throw new InvalidOperationException("The Excel file does not contain a valid workbook part. Ensure the file is a valid Excel document.");
        WorksheetPart = WorkbookPart.WorksheetParts.FirstOrDefault() ??
            throw new InvalidOperationException("The Excel file does not contain a valid worksheet part. Ensure the file is a valid Excel document.");
        SharedStrings = WorkbookPart.SharedStringTablePart?.SharedStringTable ??
            throw new InvalidOperationException("The Excel file does not contain a valid shared string part. Ensure the file is a valid Excel document.");

        var sheetDimension = GetSheetDimension();
        if (sheetDimension != null && sheetDimension.Reference is not null)
        {
            SheetRanges = SheetRange.Parse(sheetDimension.Reference.Value.AsSpan());
        }
    }

    /// <summary>
    /// Gets the WorkbookPart of the Excel document.
    /// </summary>
    public WorkbookPart WorkbookPart { get; }

    /// <summary>
    /// Gets the WorksheetPart of the Excel document.
    /// </summary>
    public WorksheetPart WorksheetPart { get; }

    /// <summary>
    /// Gets the SharedStringTable of the Excel document, which contains all shared strings used in the workbook.
    /// </summary>
    public SharedStringTable SharedStrings { get; }

    /// <summary>
    /// Gets the sheet ranges of the Excel document, which represents the dimensions of the worksheet.
    /// </summary>
    public SheetRange? SheetRanges { get; }   

    /// <summary>
    /// Retrieves the sheet dimension from the worksheet part, which indicates the range of cells that contain data in the worksheet.
    /// </summary>
    /// <returns></returns>
    private SheetDimension? GetSheetDimension()
    {
        using var xmlReader = OpenXmlReader.Create(WorksheetPart);
        while (xmlReader.Read())
        {
            if (xmlReader.ElementType == typeof(SheetDimension) && xmlReader.IsStartElement)
            {
                return xmlReader.LoadCurrentElement() as SheetDimension;
            }
        }

        return null;
    }       

    /// <summary>
    /// Disposes the ExcelImportContext, releasing any resources it holds.
    /// </summary>
    public void Dispose()
    {
        Dispose(disposing: true);
        GC.SuppressFinalize(this);
    }

    private void Dispose(bool disposing)
    {
        if (_disposed)
            return;

        if (disposing)
        {         
            _document?.Dispose();
        }

        _disposed = true;
    }
}