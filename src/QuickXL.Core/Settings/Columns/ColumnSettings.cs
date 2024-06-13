using QuickXL.Core.Styles.Columns;

namespace QuickXL.Core.Settings.Columns;

public class ColumnSettings
{
    /// <summary>
    /// If defined, column header name will be value of ColumnName.
    /// </summary>
    public string? HeaderName { get; set; }

    /// <summary>
    /// Control for related column could have empty cells or not. Defaults to true.
    /// </summary>
    public bool AllowEmptyCells { get; set; }

    public bool AutoSizeColumns { get; set; }

    public XLCellStyle HeaderStyle { get; set; }
    public XLCellStyle CellStyle { get; set; }

    public ColumnSettings()
    {
        AllowEmptyCells = true;
        HeaderStyle = new();
        CellStyle = new();
    }
}
