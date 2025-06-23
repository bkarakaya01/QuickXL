namespace QuickXL.Exporter
{
    /// <summary>
    /// Global settings for an export session (sheet name, first row index).
    /// </summary>
    public class WorkbookSettings
    {
        /// <summary>
        /// Zero‐based index of the header row.
        /// </summary>
        public int FirstRowIndex { get; set; } = 0;

        /// <summary>
        /// Name of the worksheet to create.
        /// </summary>
        public string SheetName { get; set; } = "QuickXL";
    }
}