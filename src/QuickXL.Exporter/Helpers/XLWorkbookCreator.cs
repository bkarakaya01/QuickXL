using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace QuickXL
{
    internal class XLWorkbookCreator<TDto>(ExportSettings workbookSettings, IList<TDto> excelData) where TDto : class, new()
    {
        private readonly ExportSettings WorkbookSettings = workbookSettings;
        private readonly IList<TDto> ExcelData = excelData;

        public XSSFWorkbook CreateWorkbook()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet(WorkbookSettings.SheetName);

            // Create XLSheet object.
            var xlsheet = new XLSheet<TDto>(WorkbookSettings.FirstRowIndex);

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
            var headers = XLSheetHelper<TDto>.GetHeaders();
            int columnIndex = 0;
            foreach (var header in headers)
            {
                xlsheet.AddHeader(columnIndex++, header.Value);
            }

            for (int index = 0; index < ExcelData.Count; index++)
            {
                var item = ExcelData[index];
                XLSheetHelper<TDto>.MapPocoToSheet(item, index + xlsheet.FirstRowIndex + 1, headers, xlsheet);
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
    }
}
