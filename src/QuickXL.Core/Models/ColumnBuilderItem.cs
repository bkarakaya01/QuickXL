using QuickXL.Core.Settings.Columns;

namespace QuickXL.Core.Models
{
    internal class ColumnBuilderItem<TDto> where TDto : class, new()
    {
        public string Header { get; set; }

        public Func<TDto, object> PropertySelector { get; set; }

        public ColumnSettings ColumnSettings { get; set; }
    }
}
