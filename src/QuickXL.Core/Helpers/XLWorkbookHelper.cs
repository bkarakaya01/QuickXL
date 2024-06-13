using Ardalis.GuardClauses;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using QuickXL.Core.Builders;
using QuickXL.Core.Models;
using QuickXL.Core.Models.Cells;
using QuickXL.Core.Models.Rows;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Helpers;

internal sealed class XLWorkbookHelper<TDto> where TDto : class, new()
{
    private static readonly XLWorkbookHelper<TDto> _instance = new();
    private const int InitialColumnIndex = 0;
    private XLSheet<TDto>? _xlsheet;

    private XLWorkbookHelper() { }

    public static XLWorkbookHelper<TDto> Instance => _instance;

    public XSSFWorkbook CreateWorkbook(ExportBuilder<TDto> exportBuilder, WorkbookSettings workbookSettings)
    {
        Guard.Against.Null(exportBuilder);
        Guard.Against.Null(workbookSettings);

        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet(workbookSettings.SheetName);
        _xlsheet = new XLSheet<TDto>(sheet, workbookSettings.FirstRowIndex);

        PopulateSheet(exportBuilder, workbookSettings);
        TransferDataToSheet();
        ApplyStyles(exportBuilder);

        return workbook;
    }

    private void PopulateSheet(ExportBuilder<TDto> exportBuilder, WorkbookSettings workbookSettings)
    {
        Guard.Against.Null(_xlsheet);

        Dictionary<string, string> headers = GetHeaders(exportBuilder);
        AddHeaderRow([.. headers.Keys], workbookSettings.FirstRowIndex);

        List<TDto> excelData = exportBuilder.Data;
        AddDataRows([.. headers.Values], excelData);
    }

    private static Dictionary<string, string> GetHeaders(ExportBuilder<TDto> exportBuilder)
    {
        return exportBuilder.ColumnBuilder.ColumnBuilderItems.Select(x => new KeyValuePair<string, string>(x.HeaderName, x.GetPropertyName())).ToDictionary();
    }

    private void AddHeaderRow(List<string> headerNames, int rowIndex)
    {
        _xlsheet!.AddRow(rowIndex);
        var headerRow = _xlsheet.GetRow(rowIndex);
        Guard.Against.Null(headerRow);

        for (int columnIndex = 0; columnIndex < headerNames.Count; columnIndex++)
        {
            AddCellToRow(headerRow, columnIndex, headerNames[columnIndex], isHeaderCell: true);
        }
    }

    private void AddDataRows(List<string> headerProperties, List<TDto> excelData)
    {
        for (int index = 0; index < excelData.Count; index++)
        {
            var item = excelData[index];
            int rowIndex = index + _xlsheet!.FirstRowIndex + 1;
            AddDataRow(item, headerProperties, rowIndex);
        }
    }

    private void AddDataRow(TDto item, List<string> headerProperties, int rowIndex)
    {
        _xlsheet!.AddRow(rowIndex);
        var row = _xlsheet.GetRow(rowIndex);
        Guard.Against.Null(row);
        MapPocoToSheet(item, row, headerProperties);
    }

    private void MapPocoToSheet(TDto item, XLRow<TDto> row, List<string> headerProperties)
    {
        for (int columnIndex = InitialColumnIndex; columnIndex < headerProperties.Count; columnIndex++)
        {
            var header = headerProperties[columnIndex];
            var propertyValue = GetPropertyValue(item, header);
            AddCellToRow(row, columnIndex, propertyValue);
        }
    }

    private void AddCellToRow(XLRow<TDto> row, int columnIndex, string value, bool isHeaderCell = false)
    {
        Guard.Against.Null(row);
        _xlsheet!.AddCell(row, columnIndex, value, isHeaderCell);
    }

    private static string GetPropertyValue(TDto item, string header)
    {
        return item.GetType().GetProperty(header)?.GetValue(item)?.ToString() ?? string.Empty;
    }

    private void TransferDataToSheet()
    {
        var lastRow = _xlsheet!.GetLastRow();
        for (int rowIndex = 0; rowIndex <= lastRow; rowIndex++)
        {
            var row = _xlsheet.Sheet.CreateRow(rowIndex);
            TransferRowData(row, rowIndex);
        }
    }

    private void TransferRowData(IRow row, int rowIndex)
    {
        var lastColumn = _xlsheet!.GetLastColumn(rowIndex);
        for (int columnIndex = 0; columnIndex <= lastColumn; columnIndex++)
        {
            var cellValue = _xlsheet[rowIndex, columnIndex];

            var xlCell = _xlsheet[rowIndex, columnIndex];
            Guard.Against.Null(xlCell);

            var cell = row.CreateCell(columnIndex);
            cell.SetCellValue(cellValue?.Value);

            xlCell!.Cell = cell;
        }
    }

    private void ApplyStyles(ExportBuilder<TDto> exportBuilder)
    {
        foreach (var columnBuilderItem in exportBuilder.ColumnBuilder.ColumnBuilderItems)
        {
            var column = _xlsheet!.GetColumn(columnBuilderItem.HeaderName);

            foreach (var cell in column.XLCells)
            {
                if (cell.IsHeaderCell)
                {
                    columnBuilderItem.ColumnSettings.HeaderStyle.Apply(_xlsheet, cell);
                }
                else
                    columnBuilderItem.ColumnSettings.CellStyle.Apply(_xlsheet, cell);
            }
        }

        ApplyAutoSizeColumns(exportBuilder);
    }

    private void ApplyAutoSizeColumns(ExportBuilder<TDto> exportBuilder)
    {
        var cells = _xlsheet![_xlsheet.FirstRowIndex];

        var clSettings = exportBuilder.ColumnBuilder.ColumnBuilderItems.Select(x => x.ColumnSettings);

        foreach (var columnSetting in clSettings)
        {
            if (columnSetting.AutoSizeColumns)
            {
                XLCell? cell = cells.FirstOrDefault(x => x.Value == columnSetting.HeaderName);

                Guard.Against.Null(cell);

                _xlsheet.Sheet.AutoSizeColumn(cell.ColumnIndex);
            }
        }
    }
}
