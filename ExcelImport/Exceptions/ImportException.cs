namespace ExcelImport.Exceptions;

/// <summary>
/// Represents an exception is thrown during the import process 
/// </summary>
public class ImportException : ApplicationException
{
    /// <summary>
    /// Gets the identifier of object failed to import
    /// </summary>
    public string ImportIdentifier { get; }

    /// <summary>
    /// Construct <c>ImportException</c>
    /// </summary>
    /// <param name="identifer">The identifier of object failed to import</param>
    /// <param name="message">The error message</param>
    public ImportException(string identifer, string message) : base($"[{identifer}] Import failed: {message}") => ImportIdentifier = identifer;

    /// <summary>
    /// Construct <c>ImportException</c>
    /// </summary>
    /// <param name="identifer">The identifier of object failed to import</param>
    /// <param name="innerException">The inner exception that has been thrown</param>
    public ImportException(string identifer, Exception innerException) : base($"[{identifer}] Import failed: {innerException.Message}", innerException) => ImportIdentifier = identifer;
}
