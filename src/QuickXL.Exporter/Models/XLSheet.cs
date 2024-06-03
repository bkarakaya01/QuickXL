namespace QuickXL
{
    internal sealed class XLSheet<TPoco>(int firstRowIndex = 0) : IExcelSheet where TPoco : class, IExcelPOCO, new()
    {
        private readonly IList<XLCell> _cells = [];
        public int FirstRowIndex { get; private set; } = firstRowIndex;

        /// <summary>
        /// Public getter of the specified <see cref="ExcelCell"/>.
        /// <para>
        ///     This getter will be used like a 2D Matrix array.
        /// </para>
        /// </summary>
        /// <param name="rowIndex">Specified row index.</param>
        /// <param name="columnIndex">Specified column index.</param>
        /// <returns></returns>
        internal XLCell? this[int rowIndex, int columnIndex]
        {
            get
            {
                return _cells.FirstOrDefault(x => x.RowIndex == rowIndex && x.ColumnIndex == columnIndex);
            }
        }

        /// <summary>
        /// Public getter which will get the list of <see cref="ExcelCell"/>s column data of the specified <paramref name="rowIndex"/>.
        /// </summary>
        /// <param name="rowIndex">Specified row index.</param>
        /// <returns></returns>
        internal IList<XLCell> this[int rowIndex]
        {
            get
            {
                return _cells.Where(x => x.RowIndex == rowIndex).ToList();
            }
        }

        public void AddCell(int rowIndex, int columnIndex, string value)
        {
            if (this[rowIndex, columnIndex] != null)
                throw new InvalidOperationException("Cell already exists.");

            _cells.Add(new XLCell
            {
                RowIndex = rowIndex,
                ColumnIndex = columnIndex,
                Value = value
            });
        }

        public void AddHeader(int columnIndex, string header)
        {
            AddCell(FirstRowIndex, columnIndex, header);
        }
        public int GetLastRow()
        {
            return _cells.Any() 
                ? _cells.Max(c => c.RowIndex) 
                : FirstRowIndex;
        }

        public int GetLastColumn(int rowIndex)
        {
            return _cells.Where(c => c.RowIndex == rowIndex).Any() 
                ? _cells.Where(c => c.RowIndex == rowIndex).Max(c => c.ColumnIndex) 
                : 0;
        }
    }
}
