using System.Collections.Concurrent;
using System.Linq.Expressions;
using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelImport.Converters.BuiltIn;

namespace ExcelImport.Converters;

/// <summary>
/// Base class for all converters.
/// </summary>
public abstract class TypeConverter
{
    private static readonly ConcurrentDictionary<Type, TypeConverter> _converters = new();
    private static readonly ConcurrentDictionary<Type, Func<TypeConverter>> _factoryCache = new();

    static TypeConverter()
    {
        _converters.TryAdd(typeof(int), new IntConverter());
        _converters.TryAdd(typeof(long), new LongConverter());
        _converters.TryAdd(typeof(float), new FloatConverter());
        _converters.TryAdd(typeof(double), new DoubleConverter());
        _converters.TryAdd(typeof(decimal), new DecimalConverter());
        _converters.TryAdd(typeof(bool), new BoolConverter());
        _converters.TryAdd(typeof(DateTime), new DateTimeConverter());
        _converters.TryAdd(typeof(string), new StringConverter());
    }

    /// <summary>
    /// Creates a the specified type converter
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentException"></exception>
    public static TypeConverter CreateConverter(Type type)
    {
        if (_converters.TryGetValue(type, out var converter))
        {
            return converter;
        }

        ConstructorInfo? ctor = type.GetConstructor(Type.EmptyTypes);
        if (!typeof(TypeConverter).IsAssignableFrom(type) || ctor == null || !ctor.IsPublic)
        {
            throw new ArgumentException($"Type {type} is not a valid converter type");
        }

        var creator = _factoryCache.GetOrAdd(type, CreateFactory);
        var convertedType = creator();
        _converters.TryAdd(type, convertedType);

        return convertedType;
    }

    /// <summary>
    /// Creates a type converter for the specified type.
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    public static TypeConverter CreateDefaultTypeConverter(Type type)
    {
        var genericType = typeof(DefaultTypeConverter<>).MakeGenericType(type);

        var creator = _factoryCache.GetOrAdd(genericType, CreateFactory);
        return creator();
    }

    /// <summary>
    /// Creates a factory method for the specified type converter type.
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    private static Func<TypeConverter> CreateFactory(Type type)
    {
        var ctor = type.GetConstructor(Type.EmptyTypes);
        if (ctor == null)
            throw new InvalidOperationException($"No parameterless constructor found for {type}");

        var newExpr = Expression.New(ctor);
        var castExpr = Expression.Convert(newExpr, typeof(TypeConverter));

        return Expression.Lambda<Func<TypeConverter>>(castExpr).Compile();
    }
       
    /// <summary>
    /// Gets the type converter for the specified type.
    /// <para If the converter does not exist, it will create a new one and cache it.</para>
    /// </summary>
    /// <param name="type"></param>
    /// <returns></returns>
    public static TypeConverter GetConverterOfType(Type type)
    {
        if (_converters.TryGetValue(type, out var converter))
        {
            return converter;
        }

        if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            var underlyingType = Nullable.GetUnderlyingType(type);
            if (underlyingType != null && _converters.TryGetValue(underlyingType, out converter))
            {
                return converter;
            }
        }

        var converterType = CreateDefaultTypeConverter(type);

        if (converterType != null)
        {
            _converters.TryAdd(type, converterType);
        }

        return converterType ?? new DefaultTypeConverter<object>();
    }

    /// <summary>
    /// The cell type of the converter.
    /// </summary>
    public abstract CellValues CellType { get; }

    /// <summary>
    /// The style index of the cell.
    /// </summary>
    public abstract uint? StyleIndex { get; }

    /// <summary>
    /// Converts the value to the target type.
    /// </summary>
    /// <param name="value"></param>
    /// <param name="throwOnConvertError"></param>
    /// <returns></returns>
    public abstract object? Convert(object value, bool throwOnConvertError = false);

    /// <summary>
    /// Converts the value to the string representation of the target type.
    /// </summary>
    /// <param name="value"></param>
    /// <param name="throwOnConvertError"></param>
    /// <returns></returns>
    public abstract CellValue? ConvertToCellValue(object value, bool throwOnConvertError = false);
}