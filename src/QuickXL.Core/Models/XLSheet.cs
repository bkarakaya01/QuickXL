using Ardalis.GuardClauses;
using QuickXL.Core.Models.Cells;
using QuickXL.Core.Models.Columns;
using QuickXL.Core.Models.Rows;

namespace QuickXL.Core.Models;

internal sealed class XLSheet<TDto>(int firstRowIndex = 0) where TDto : class, new()
{
    private readonly IList<XLRow<TDto>> _rows = [];
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
            return _rows.FirstOrDefault(x => x.Index == rowIndex)?.Cells.FirstOrDefault(x => x.ColumnIndex == columnIndex);
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
            return _rows.FirstOrDefault(x => x.Index == rowIndex)?.Cells.ToList() ?? [];
        }
    }

    internal XLColumn GetColumn(int columnIndex)
    {
        XLRow<TDto>? headerRow = _rows.FirstOrDefault(x => x.IsHeaderRow);
        Guard.Against.Null(headerRow);

        XLCell? headerCell = headerRow.GetCell(columnIndex);
        Guard.Against.Null(headerCell);

        var column = new XLColumn()
        {
            Index = columnIndex,
            HeaderName = headerCell.Value ?? string.Empty,
        };

        var rows = _rows.OrderBy(x => x.Index).ToList();

        foreach (var row in rows)
        {
            var cell = this[row.Index, columnIndex];

            if (cell != null)
                column.Add(cell);
        }
        return column;
    }

    internal void AddRow(int rowIndex)
    {
        _rows.Add(new()
        {
            Index = rowIndex,
            IsHeaderRow = rowIndex == FirstRowIndex
        });
    }

    internal void AddCell(XLRow<TDto> row, int columnIndex, string value, bool isHeaderCell = false)
    {
        if (isHeaderCell && row.Index != FirstRowIndex)
            throw new InvalidOperationException($"Header cells must be defined at {FirstRowIndex}");

        if (this[row.Index, columnIndex] != null)
            throw new InvalidOperationException("Cell already exists.");

        row.Cells.Add(new XLCell
        {
            Value = value,
            IsHeaderCell = isHeaderCell
        });
    }

    internal XLRow<TDto>? GetRow(int rowIndex) => _rows.FirstOrDefault(x => x.Index == rowIndex);

    internal int GetLastRow()
    {
        return _rows.Max(x => x.Index);
    }

    internal int GetLastColumn(int rowIndex)
    {
        XLRow<TDto>? row = GetRow(rowIndex);

        Guard.Against.Null(row);

        return row.Cells.Max(x => x.ColumnIndex);
    }
}
