using QuickXL.Core.Models.Columns;

namespace QuickXL.Core.Models.Rows
{
    internal class XLRow<TDto>
    {
        protected internal int Index { get; set; }
        protected internal bool IsHeaderRow { get; set; }

        protected internal IList<XLColumn<TDto>> Columns { get; set; }
        public XLRow()
        {
            Columns = [];
        }

        public XLColumn<TDto>? GetColumn(int columnIndex) => Columns.FirstOrDefault(x => x.Index == columnIndex);
    }
}
