namespace ExcelImport.Converters;

/// <summary>
/// Interface for all converters.
/// </summary>
/// <typeparam name="T"></typeparam>
/// <typeparam name="Source"></typeparam>
public interface IConverter<out T, in Source>
{
    /// <summary>
    /// Converts a value from the source type to the target type.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public T Convert(Source value);
}