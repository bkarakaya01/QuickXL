using QuickXL.Core.Models.Columns;

namespace QuickXL.Core.Settings.Columns;

public class ColumnSettings
{
    /// <summary>
    /// Control for related column could have empty cells or not. Defaults to true.
    /// </summary>
    public bool AllowEmptyCells { get; set; }

    public ColumnSettings()
    {
        AllowEmptyCells = true;
    }
}
