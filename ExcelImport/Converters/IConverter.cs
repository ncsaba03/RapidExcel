namespace ExcelImport.Converters;

/// <summary>
/// Interface for all converters.
/// </summary>
/// <typeparam name="TTarget">The target type to convert to</typeparam>
/// <typeparam name="TSource">The source type to convert from</typeparam>
public interface IConverter<out TTarget, in TSource>
{
    /// <summary>
    /// Converts a value from the source type to the target type.
    /// </summary>
    /// <param name="value"></param>
    /// <returns></returns>
    public TTarget? Convert(TSource value);
}