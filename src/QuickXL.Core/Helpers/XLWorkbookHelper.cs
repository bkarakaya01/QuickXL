using Ardalis.GuardClauses;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using QuickXL.Core.Builders;
using QuickXL.Core.Models;
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
        _xlsheet = new XLSheet<TDto>(workbookSettings.FirstRowIndex);

        PopulateSheet(exportBuilder, workbookSettings);
        TransferDataToSheet(sheet);

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
            AddColumnToRow(headerRow, columnIndex, headerNames[columnIndex]);
        }
    }

    private void AddColumnToRow(XLRow<TDto> row, int columnIndex, string header)
    {
        _xlsheet!.AddColumn(row, columnIndex, header);
        var column = row.GetColumn(columnIndex);
        Guard.Against.Null(column);
        _xlsheet.AddCell(column, header);
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

    private static string GetPropertyValue(TDto item, string header)
    {
        return item.GetType().GetProperty(header)?.GetValue(item)?.ToString() ?? string.Empty;
    }

    private void AddCellToRow(XLRow<TDto> row, int columnIndex, string propertyValue)
    {
        _xlsheet!.AddColumn(row, columnIndex, propertyValue);
        var column = row.GetColumn(columnIndex);
        Guard.Against.Null(column);
        _xlsheet.AddCell(column, propertyValue);
    }

    private void TransferDataToSheet(ISheet sheet)
    {
        var lastRow = _xlsheet!.GetLastRow();
        for (int rowIndex = 0; rowIndex <= lastRow; rowIndex++)
        {
            var row = sheet.CreateRow(rowIndex);
            TransferRowData(row, rowIndex);
        }
    }

    private void TransferRowData(IRow row, int rowIndex)
    {
        var lastColumn = _xlsheet!.GetLastColumn(rowIndex);
        for (int columnIndex = 0; columnIndex <= lastColumn; columnIndex++)
        {
            var cellValue = _xlsheet[rowIndex, columnIndex];
            var cell = row.CreateCell(columnIndex);
            cell.SetCellValue(cellValue?.Value);
        }
    }
}
