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
        /// Builds a setter action that parses the cell text into the property's type and assigns it.
        /// Supports <see cref="int"/>, <see cref="double"/>, <see cref="bool"/>, <see cref="DateTime"/>,
        /// and defaults to string for other types.
        /// </summary>
        /// <param name="property">The property to set on the DTO.</param>
        /// <returns>An action that takes the DTO instance and raw cell text.</returns>
        internal static Action<object, string> BuildSetter(PropertyInfo property)
        {
            return (dto, text) =>
            {
                object value = text;

                if (property.PropertyType == typeof(int) && int.TryParse(text, out var iv))
                    value = iv;
                else if (property.PropertyType == typeof(double) && double.TryParse(text, out var dv))
                    value = dv;
                else if (property.PropertyType == typeof(bool) && bool.TryParse(text, out var bv))
                    value = bv;
                else if (property.PropertyType == typeof(DateTime) && DateTime.TryParse(text, out var dt))
                    value = dt;

                property.SetValue(dto, value);
            };
        }
    }
}
