namespace QuickXL.Importer
{
    /// <summary>
    /// Configuration settings for importing Excel data into POCOs.
    /// </summary>
    public class ImportSettings
    {
        /// <summary>
        /// When <c>true</c>, automatically map columns to properties
        /// using <see cref="XLColumnAttribute"/> on POCO properties.
        /// </summary>
        public bool UseAttributes { get; set; } = true;

        /// <summary>
        /// Options for selecting which worksheet to import (by name or index).
        /// </summary>
        public SheetSettings? SheetSettings { get; set; }

        /// <summary>
        /// Options for locating the header row (manual index or auto-detection).
        /// </summary>
        public HeaderRowSettings? HeaderRowSettings { get; set; }

        /// <summary>
        /// Additional mappings specified manually (column index to setter).
        /// </summary>
        internal List<Mapping> ManualMappings { get; } = new List<Mapping>();

        /// <summary>
        /// Initializes a new <see cref="ImportSettings"/>
        /// with default sheet and header row settings.
        /// </summary>
        public ImportSettings()
        {
            SheetSettings = new SheetSettings();
            HeaderRowSettings = new HeaderRowSettings();
        }
    }
}
