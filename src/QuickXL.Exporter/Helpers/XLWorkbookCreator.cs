using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace QuickXL
{
    internal class XLWorkbookCreator<TPoco>(WorkbookSettings workbookSettings, IList<TPoco> excelData) where TPoco : class, IExcelPOCO, new()
    {
        private readonly WorkbookSettings WorkbookSettings = workbookSettings;
        private readonly IList<TPoco> ExcelData = excelData;

        public XSSFWorkbook CreateWorkbook()
        {
            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet(WorkbookSettings.SheetName);

            // Create XLSheet object.
            var xlsheet = new XLSheet<TPoco>(WorkbookSettings.FirstRowIndex);

            //  Verileri işleyip XLSheet nesnesine ekleyin
            PopulateSheet(xlsheet);

            // Transfer data defined under XLSheet into NPOI sheet.
            XLWorkbookCreator<TPoco>.TransferDataToSheet(sheet, xlsheet);

            return workbook;
        }

        /// <summary>
        /// Populate <see cref="XLSheet{TPoco}"/> data.
        /// </summary>
        /// <param name="xlsheet"></param>
        private void PopulateSheet(XLSheet<TPoco> xlsheet)
        {
            var headers = XLSheetHelper<TPoco>.GetHeaders();
            int columnIndex = 0;
            foreach (var header in headers)
            {
                xlsheet.AddHeader(columnIndex++, header.Value);
            }

            for (int index = 0; index < ExcelData.Count; index++)
            {
                var item = ExcelData[index];
                XLSheetHelper<TPoco>.MapPocoToSheet(item, index + xlsheet.FirstRowIndex + 1, headers, xlsheet);
            }
        }

        /// <summary>
        /// Transfer <see cref="XLSheet{TPoco}"/> data into <see cref="ISheet"/>.
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="xlsheet"></param>
        private static void TransferDataToSheet(ISheet sheet, XLSheet<TPoco> xlsheet)
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
