namespace QuickXL.Importer
{
    /// <summary>
    /// Represents a mapping between a column index in the worksheet and
    /// a setter action that assigns the cell text to a DTO property.
    /// </summary>
    internal class Mapping
    {
        /// <summary>
        /// Zero-based index of the column in the worksheet.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// Action that assigns the raw cell text to the target DTO's property.
        /// </summary>
        public Action<object, string> Setter { get; set; } = (_, __) => { };
    }
}
