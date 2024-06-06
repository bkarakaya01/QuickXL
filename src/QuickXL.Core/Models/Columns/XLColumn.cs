using QuickXL.Core.Models.Cells;
using QuickXL.Core.Settings.Columns;

namespace QuickXL.Core.Models.Columns;

internal class XLColumn<TDto>
{
    protected required internal int Index { get; set; }
    protected required internal string HeaderName { get; set; }
    protected internal Func<TDto, object> PropertySelector { get; set; }
    protected internal ColumnSettings ColumnSettings { get; set; }    
    protected internal XLCell XLCell { get; set; }

    //public XLColumn(int index, string name, Func<TDto, object> propertySelector, ColumnSettings columnSettings)
    //{
    //    Index = index;
    //    HeaderName = name;
    //    PropertySelector = propertySelector;
    //    ColumnSettings = columnSettings;
    //}    
}
