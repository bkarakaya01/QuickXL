using QuickXL.Core.Models.Cells;
using QuickXL.Core.Settings.Columns;

namespace QuickXL.Core.Models.Columns;

internal class XLColumn<TDto>
{
    protected required internal int Index { get; set; }
    protected required internal string HeaderName { get; set; }
    protected internal List<XLCell> XLCells { get; set; }
}
