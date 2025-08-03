using System.Collections.Concurrent;
using System.Reflection;
using ExcelImport.Attributes;
using ExcelImport.Converters;

namespace ExcelImport
{
    internal static class PropertyCache
    {
        private static readonly ConcurrentDictionary<Type, List<PropertyImportInfo>> _propertyCache = new();

        /// <summary>
        /// Gets the properties of the type with the ImportAttribute.
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static List<PropertyImportInfo> GetCachedProperties(Type type)
        {
            return _propertyCache.GetOrAdd(type, t =>
              [.. t.GetProperties(BindingFlags.Public | BindingFlags.Instance)
             .Where(p => p.IsDefined(typeof(ExcelColumnAttribute)))
             .Select(p => new PropertyImportInfo
             {
                 Property = p,
                 ColumnIdentifier = p.GetCustomAttribute<ExcelColumnAttribute>()!.Name,
                 TypeConverter = p.GetCustomAttribute<ExcelColumnAttribute>()?.CreateConverter() ?? TypeConverter.GetConverterOfType(p.PropertyType),
                 Position = p.GetCustomAttribute<ExcelColumnAttribute>()?.Position ?? int.MaxValue,
                 Required = p.GetCustomAttribute<ExcelColumnAttribute>()?.Required ?? false

             })
             .OrderBy(t => t.Position)]);
        }
    }
}
