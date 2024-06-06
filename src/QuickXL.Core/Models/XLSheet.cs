using QuickXL.Core.Models.Cells;

namespace QuickXL.Core.Models;

internal sealed class XLSheet<TDto>(int firstRowIndex = 0) where TDto : class, new()
{
    private readonly IList<XLCell> _cells = [];
    internal int FirstRowIndex { get; private set; } = firstRowIndex;

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

    internal void AddCell(int rowIndex, int columnIndex, string value)
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

    internal void AddHeader(int columnIndex, string header)
    {
        AddCell(FirstRowIndex, columnIndex, header);
    }
    internal int GetLastRow()
    {
        return _cells.Any()
            ? _cells.Max(c => c.RowIndex)
            : FirstRowIndex;
    }

    internal int GetLastColumn(int rowIndex)
    {
        return _cells.Where(c => c.RowIndex == rowIndex).Any()
            ? _cells.Where(c => c.RowIndex == rowIndex).Max(c => c.ColumnIndex)
            : 0;
    }
}
