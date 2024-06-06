using Ardalis.GuardClauses;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using QuickXL.Core.Builders;
using QuickXL.Core.Models;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Helpers;

internal sealed class XLWorkbookHelper<TDto> where TDto : class, new()
{
    private static readonly XLWorkbookHelper<TDto> _instance = new();

    private XLWorkbookHelper() { }

    public static XLWorkbookHelper<TDto> Instance => _instance;

    public XSSFWorkbook CreateWorkbook(ExportBuilder<TDto> exportBuilder, WorkbookSettings workbookSettings)
    {
        Guard.Against.Null(exportBuilder);
        Guard.Against.Null(exportBuilder.ColumnBuilder);
        Guard.Against.Null(exportBuilder.ColumnBuilder.Columns);
        Guard.Against.Null(workbookSettings);

        var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet(workbookSettings.SheetName);

        var xlsheet = new XLSheet<TDto>(workbookSettings.FirstRowIndex);

        XLWorkbookHelper<TDto>.Instance.PopulateSheet(xlsheet, exportBuilder);

        // Transfer data defined under XLSheet into NPOI sheet.
        XLWorkbookHelper<TDto>.Instance.TransferDataToSheet(sheet, xlsheet);

        return workbook;
    }

    /// <summary>
    /// Populate <see cref="XLSheet{TPoco}"/> data.
    /// </summary>
    /// <param name="xlsheet"></param>
    private void PopulateSheet(XLSheet<TDto> xlsheet, ExportBuilder<TDto> exportBuilder)
    {
        Guard.Against.Null(xlsheet);

        IList<string> headers = exportBuilder.ColumnBuilder!.Columns!.Select(x => x.Key).ToList();

        int columnIndex = 0;
        foreach (var header in headers)
        {
            xlsheet.AddHeader(columnIndex++, header);
        }

        List<TDto> excelData = exportBuilder.Data;

        for (int index = 0; index < excelData.Count; index++)
        {
            var item = excelData[index];
            int rowIndex = index + xlsheet.FirstRowIndex + 1;
            MapPocoToSheet(item, rowIndex, headers, xlsheet);
        }
    }

    /// <summary>
    /// Transfer <see cref="XLSheet{TPoco}"/> data into <see cref="ISheet"/>.
    /// </summary>
    /// <param name="sheet"></param>
    /// <param name="xlsheet"></param>
    private void TransferDataToSheet(ISheet sheet, XLSheet<TDto> xlsheet)
    {
        for (int i = 0; i <= xlsheet.GetLastRow(); i++)
        {
            var row = sheet.CreateRow(i);
            for (int j = 0; j <= xlsheet.GetLastColumn(i); j++)
            {
                var cellValue = xlsheet[i, j];
                var cell = row.CreateCell(j);
                cell.SetCellValue(cellValue?.Value);
            }
        }
    }

    private static void MapPocoToSheet(TDto item, int rowIndex, IList<string> headers, XLSheet<TDto> sheet)
    {
        int columnIndex = 0;
        foreach (var header in headers)
        {
            var propertyValue = item.GetType().GetProperty(header)?.GetValue(item)?.ToString() ?? string.Empty;
            sheet.AddCell(rowIndex, columnIndex++, propertyValue);
        }
    }
}
