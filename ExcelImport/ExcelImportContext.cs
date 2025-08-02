using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Spreadsheet;

public sealed class ExcelImportContext : IDisposable
{
    private readonly SpreadsheetDocument _document;

    public WorkbookPart WorkbookPart { get; }
    public WorksheetPart WorksheetPart { get; }
    public SharedStringTable SharedStrings { get; }
    public SheetRange? SheetRanges { get; }

    public ExcelImportContext(string path)
    {
        _document = SpreadsheetDocument.Open(path, false);
        WorkbookPart = _document.WorkbookPart ?? throw new InvalidOperationException("The Excel file does not contain a valid workbook part. Ensure the file is a valid Excel document.");
        WorksheetPart = WorkbookPart.WorksheetParts.FirstOrDefault() ?? throw new InvalidOperationException("The Excel file does not contain a valid worksheet part. Ensure the file is a valid Excel document.");
        SharedStrings = WorkbookPart.SharedStringTablePart?.SharedStringTable ?? throw new InvalidOperationException("The Excel file does not contain a valid shared string part. Ensure the file is a valid Excel document.");
        var sheetDimension = GetSheetDimension();   
        if (sheetDimension != null && sheetDimension.Reference is not null)
        {
            SheetRanges = SheetRange.Parse(sheetDimension.Reference.Value.AsSpan());
        }
    }

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

    private bool _disposed = false;

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