using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelImport.Converters
{    
    /// <summary>
    /// Base class for all converters.
    /// </summary>
    /// <typeparam name="Type"></typeparam>
    /// <typeparam name="Source"></typeparam>
    public abstract class TypeConverter<Type, Source> : TypeConverter, IConverter<Type, Source>
    {
        public override CellValues CellType => CellValues.String;
        public override uint? StyleIndex => null;

        /// <summary>
        /// Converts the value to the target type.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="throwOnConvertError"></param>
        /// <returns></returns>
        /// <exception cref="InvalidCastException"></exception>
        public override object? Convert(object value, bool throwOnConvertError = false)
        {
            var converted = Convert((Source)value);

            if (converted is null && throwOnConvertError)
            {
                throw new InvalidCastException($"Cannot convert {value} to {typeof(Type)}");
            }

            return converted;
        }

        public override CellValue? ConvertToCellValue(object value, bool throwOnConvertError = false)
        {
            var converted = ConvertToCellValue((Type)value);

            if (converted is null && throwOnConvertError)
            {
                throw new InvalidCastException($"Cannot convert {value} to {typeof(string)}");
            }

            return converted;
        }



        /// <summary>
        /// Converts the value to the target type.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public abstract Type Convert(Source value);

        /// <summary>
        /// Converts the value to the string representation of the target type.
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public abstract CellValue? ConvertToCellValue(Type value);
    }   
}