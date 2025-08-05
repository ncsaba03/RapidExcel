using System.Reflection;
using ExcelImport.Converters;

namespace ExcelImport.Attributes;

/// <summary>
/// Attribute for import
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    /// <summary>
    /// Constructs <c>ExcelColumnAttribute</c>
    /// </summary>
    /// <param name="name">The name of the attribute</param>
    /// <param name="position">The position of the data</param>
    /// <param name="typeConverter">The type to convert to</param>
    /// <param name="required">Whether the property is mandatory</param>
    public ExcelColumnAttribute(string name, int position = -1, Type? typeConverter = null, bool required = false)
    {
        Name = name;
        Position = position;
        ConverterType = typeConverter;
        Required = required;
    }

    /// <summary>
    /// Gets or sets the name
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Gets or sets the position
    /// </summary>
    public int Position { get; set; }

    /// <summary>
    /// Gets or sets the type converter
    /// </summary>
    public Type? ConverterType { get; set; }

    /// <summary>
    /// Gets or sets that the property is required
    /// </summary>
    public bool Required { get; set; }

    public TypeConverter? CreateConverter()
    {
        if (ConverterType != null)
        {
            ConstructorInfo? ctor = ConverterType.GetConstructor(Type.EmptyTypes);
            if (!typeof(TypeConverter).IsAssignableFrom(ConverterType) || ctor == null || !ctor.IsPublic)
            {
                throw new ArgumentException($"Type {ConverterType} is not a valid converter type");
            }

            return (TypeConverter?)Activator.CreateInstance(ConverterType);
        }

        return null;
    }

}
