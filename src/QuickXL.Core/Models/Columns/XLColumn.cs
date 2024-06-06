using QuickXL.Core.Models.Cells;
using QuickXL.Core.Settings.Columns;

namespace QuickXL.Core.Models.Columns;

public record XLColumn<TDto>
{
    protected internal int Index { get; set; }
    protected internal string Name { get; set; }
    protected internal Func<TDto, object> PropertySelector { get; set; }
    protected internal ColumnSettings ColumnSettings { get; set; }

    internal XLHeaderCell? XLHeaderCell { get; set; }
    internal IList<XLCell> XLCells { get; set; }

    public XLColumn(int index, string name, Func<TDto, object> propertySelector, ColumnSettings columnSettings)
    {
        Index = index;
        Name = name;
        PropertySelector = propertySelector;
        ColumnSettings = columnSettings;

        XLCells = [];
    }
}
