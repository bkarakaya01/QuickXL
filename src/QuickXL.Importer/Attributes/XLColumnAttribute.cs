namespace QuickXL.Importer
{
    /// <summary>
    /// Specifies the Excel column header that a POCO property maps to when importing data.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public sealed class XLColumnAttribute : Attribute
    {
        /// <summary>
        /// Gets the header text of the Excel column that this property is bound to.
        /// </summary>
        public string HeaderName { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="XLColumnAttribute"/> class
        /// with the specified column header name.
        /// </summary>
        /// <param name="headerName">The header text of the Excel column.</param>
        /// <exception cref="ArgumentException">
        /// Thrown if <paramref name="headerName"/> is null, empty, or whitespace.
        /// </exception>
        public XLColumnAttribute(string headerName)
        {
            if (string.IsNullOrWhiteSpace(headerName))
                throw new ArgumentException("Header name must not be null or whitespace.", nameof(headerName));

            HeaderName = headerName;
        }
    }
}
