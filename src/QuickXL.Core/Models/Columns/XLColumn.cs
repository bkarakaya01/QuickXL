using QuickXL.Core.Models.Cells;

namespace QuickXL.Core.Models.Columns;

internal class XLColumn
{
    protected required internal int Index { get; set; }
    protected required internal string HeaderName { get; set; }
    protected internal List<XLCell> XLCells { get; set; }

    public XLColumn()
    {
        XLCells = [];
    }

    public void Add(XLCell cell) => XLCells.Add(cell);
}
