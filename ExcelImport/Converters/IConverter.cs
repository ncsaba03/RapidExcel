namespace ExcelImport.Converters
{
    /// <summary>
    /// Interface for all converters.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    /// <typeparam name="Source"></typeparam>
    public interface IConverter<out T, in Source>
    {
        public T Convert(Source value);
    }
}