using System.Reflection;

namespace QuickXL.Importer.Helpers
{
    /// <summary>
    /// Provides helper methods to build property mappings between Excel columns and DTO properties.
    /// </summary>
    internal static class MappingHelper
    {
        /// <summary>
        /// Scans <typeparamref name="TDto"/> public instance properties for <see cref="XLColumnAttribute"/>
        /// and generates a mapping of column index to setter action for each attributed property.
        /// </summary>
        /// <typeparam name="TDto">The target POCO type being imported.</typeparam>
        /// <param name="headerNames">Sequence of header cell text values from the worksheet.</param>
        /// <returns>A sequence of <see cref="Mapping"/> instances linking column indices to property setters.</returns>
        internal static IEnumerable<Mapping> BuildAttributeMappings<TDto>(
            IEnumerable<string> headerNames)
            where TDto : class, new()
        {
            var props = typeof(TDto)
                .GetProperties(BindingFlags.Public | BindingFlags.Instance);

            foreach (var prop in props)
            {
                var attr = prop.GetCustomAttribute<XLColumnAttribute>();
                if (attr == null)
                    continue;

                int index = headerNames
                    .Select((h, i) => new { Header = h, Index = i })
                    .FirstOrDefault(x =>
                        string.Equals(x.Header, attr.HeaderName, StringComparison.OrdinalIgnoreCase))
                    ?.Index ?? -1;

                if (index < 0)
                    continue;

                yield return new Mapping
                {
                    Index = index,
                    Setter = BuildSetter(prop)
                };
            }
        }

        /// <summary>
        /// Builds a setter that parses the cell text into the target property’s type:
        /// on success assigns the parsed value; on failure assigns default(T).
        /// </summary>
        internal static Action<object, string> BuildSetter(PropertyInfo property)
        {
            return (dto, text) =>
            {
                Type t = property.PropertyType;

                if (t == typeof(string))
                {
                    // Always assign strings
                    property.SetValue(dto, text);
                    return;
                }

                if (t == typeof(int))
                {
                    if (int.TryParse(text, out var iv))
                        property.SetValue(dto, iv);
                    else
                        property.SetValue(dto, default(int)); // 0
                    return;
                }

                if (t == typeof(double))
                {
                    if (double.TryParse(text, out var dv))
                        property.SetValue(dto, dv);
                    else
                        property.SetValue(dto, default(double)); // 0.0
                    return;
                }

                if (t == typeof(bool))
                {
                    if (bool.TryParse(text, out var bv))
                        property.SetValue(dto, bv);
                    else
                        property.SetValue(dto, default(bool)); // false
                    return;
                }

                if (t == typeof(DateTime))
                {
                    if (DateTime.TryParse(text, out var dt))
                        property.SetValue(dto, dt);
                    else
                        property.SetValue(dto, default(DateTime)); // DateTime.MinValue
                    return;
                }

                // Other types: do nothing or consider throwing
            };
        }
    }
}
