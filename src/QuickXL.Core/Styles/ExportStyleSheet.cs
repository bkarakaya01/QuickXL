using NPOI.SS.UserModel;
using QuickXL.Core.Models.Cells;
using QuickXL.Core.Styles.Columns;

namespace QuickXL.Core.Styles;

public class ExportStyleSheet
{
    /// <summary>
    /// Decides whether columns should auto sized or not.
    /// </summary>
    public bool AutoSizeColumns { get; set; }

    /// <summary>
    /// Settings class to manage the header style.
    /// </summary>
    public XLHeaderStyle XLHeaderStyle { get; set; }

    /// <summary>
    /// Settings class to manage the value cell style.
    /// </summary>
    public XLCellStyle XLCellStyle { get; set; }

    public ExportStyleSheet()
    {
        AutoSizeColumns = false;
        XLHeaderStyle = new();
        XLCellStyle = new();
    }
   
    internal void Apply(ISheet excelSheet, XLCell cell)
    {
        XLCellStyle.Apply(excelSheet, cell);
    }

    internal void ApplyAutoSizeColumns(ISheet excelSheet, int maxColumnCount)
    {
        if (AutoSizeColumns)
            for (int columnIndex = 0; columnIndex < maxColumnCount; columnIndex++)
                excelSheet.AutoSizeColumn(columnIndex);
    }
}
