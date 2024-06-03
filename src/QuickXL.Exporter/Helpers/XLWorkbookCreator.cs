using Ardalis.GuardClauses;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace QuickXL
{
    internal class XLWorkbookCreator<TDto>(Exporter<TDto> exporter) where TDto : class, new()
    {
        private readonly Exporter<TDto> Exporter = exporter;        

        public XSSFWorkbook CreateWorkbook()
        {
            Guard.Against.Null(Exporter);

            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet(Exporter.WorkbookSettings.SheetName);

            // Create XLSheet object.
            var xlsheet = new XLSheet<TDto>(Exporter.WorkbookSettings.FirstRowIndex);

            //  Verileri işleyip XLSheet nesnesine ekleyin
            PopulateSheet(xlsheet);

            // Transfer data defined under XLSheet into NPOI sheet.
            XLWorkbookCreator<TDto>.TransferDataToSheet(sheet, xlsheet);

            return workbook;
        }

        /// <summary>
        /// Populate <see cref="XLSheet{TPoco}"/> data.
        /// </summary>
        /// <param name="xlsheet"></param>
        private void PopulateSheet(XLSheet<TDto> xlsheet)
        {
            Guard.Against.Null(xlsheet);

            IList<string> headers = Exporter.ExportBuilder!.HeaderPropertySelectors.Select(x => x.Key).ToList();

            int columnIndex = 0;
            foreach (var header in headers)
            {
                xlsheet.AddHeader(columnIndex++, header);
            }

            List<TDto> excelData = Exporter.ExportBuilder!.Data;

            for (int index = 0; index < excelData.Count; index++)
            {
                var item = excelData[index];
                MapPocoToSheet(item, index + xlsheet.FirstRowIndex + 1, headers, xlsheet);
            }
        }

        /// <summary>
        /// Transfer <see cref="XLSheet{TPoco}"/> data into <see cref="ISheet"/>.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="xlsheet"></param>
        private static void TransferDataToSheet(ISheet sheet, XLSheet<TDto> xlsheet)
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
}
