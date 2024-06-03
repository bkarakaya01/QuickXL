namespace QuickXL
{
    /// <summary>
    /// An <see cref="Attribute"/> which will be used on DTO classes which are marked as an <see cref="IExcelPOCO"/>.
    /// 
    /// <para>
    ///     This attribute should be used to track the header information of an excel file. Especially is a standart for the export operations.
    /// </para>
    /// </summary>
    [AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = true)]
    public class XLHeader(string headerName) : Attribute
    {

        /// <summary>
        /// Specified header name.
        /// </summary>
        public string HeaderName { get; set; } = headerName;
    }
}
