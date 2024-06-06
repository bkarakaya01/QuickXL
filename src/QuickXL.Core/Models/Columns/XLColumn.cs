using QuickXL.Core.Settings.Columns;

namespace QuickXL.Core.Models.Columns;

public record XLColumn<TDto>(string Name, Func<TDto, object> PropertySelector, ColumnSettings? ColumnSettings = null);
