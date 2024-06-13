using QuickXL.Core.Models.Cells;

namespace QuickXL.Core.Models.Rows
{
    internal class XLRow<TDto>
    {
        protected internal int Index { get; set; }
        protected internal bool IsHeaderRow { get; set; }

        protected internal IList<XLCell> Cells { get; set; }
        public XLRow()
        {
            Cells = [];
        }

        public XLCell? GetCell(int columnIndex) => Cells.FirstOrDefault(x => x.ColumnIndex == columnIndex);
    }
}
