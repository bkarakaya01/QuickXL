namespace QuickXL.Importer
{
    /// <summary>
    /// Configure which worksheet to import, by name or index.
    /// </summary>
    public class SheetSettings
    {
        /// <summary>
        /// Zero-based sheet index to import if <see cref="SheetName"/> is null.
        /// </summary>
        public int SheetIndex { get; set; } = 0;

        /// <summary>
        /// Name of the worksheet to import.  If specified, takes precedence over <see cref="SheetIndex"/>.
        /// </summary>
        public string? SheetName { get; set; }
    }
}
