using NPOI.SS.UserModel;

namespace QuickXL
{
    // <summary>
    /// Object as a reference to a cell in an excel file.
    /// </summary>
    internal record XLCell : IExcelUnit
    {
        /// <summary>
        /// Cell value.
        /// </summary>
        public string? Value { get; set; }

        /// <summary>
        /// Column index of the current cell.
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// High level representation of a cell.
        /// </summary>
        public ICell? Cell { get; set; }

        /// <summary>
        /// Row index of the current cell.
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// High level representation of a row.
        /// </summary>
        public IRow? Row { get; set; }
    }
}
