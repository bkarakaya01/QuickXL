namespace QuickXL.Core.Models.Cells;

internal record XLHeaderCell : XLCell
{
    public bool AllowEmptyCells { get; set; }
}
