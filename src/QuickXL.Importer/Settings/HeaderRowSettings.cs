namespace QuickXL.Importer
{
    /// <summary>
    /// Configure options for detecting and selecting the header row when importing.
    /// </summary>
    public class HeaderRowSettings
    {
        /// <summary>
        /// Zero-based index of the header row.  When set, automatic detection is skipped.
        /// </summary>
        public int? StartsAt { get; set; }
    }
}
