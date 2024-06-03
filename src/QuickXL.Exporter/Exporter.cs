using NPOI.XSSF.UserModel;

namespace QuickXL
{
    public sealed class Exporter
    {

        /// <summary>
        /// Creates a <see cref="NPOI.XSSF.UserModel.XSSFWorkbook"/> to export Excel File as a <see cref="Stream"/>.
        /// </summary>
        /// <typeparam name="TPoco">Data transfer object which has to be marked as <see cref="IExcelPOCO"/> as a standart.</typeparam>
        /// <param name="workbookSettings"><see cref="WorkbookSettings"/> will be used to store detailed information.</param>
        /// <param name="excelData">List of data transfer objects which will be used to create the excel file.</param>
        /// <returns></returns>
        public static XLResult Export<TPoco>(WorkbookSettings workbookSettings, List<TPoco> excelData) where TPoco : class, IExcelPOCO, new()
        {
            var memory = new MemoryStream();

            using (var fs = new MemoryStream())
            {
                var workbookCreator = new XLWorkbookCreator<TPoco>(workbookSettings, excelData);

                XSSFWorkbook workbook = workbookCreator.CreateWorkbook();

                workbook.Write(fs);

                var byteArray = fs.ToArray();
                memory.Write(byteArray);
            }
            memory.Position = 0;

            return new XLResult()
            {
                Success = true,
                Data = memory
            };
        }

        /// <summary>
        /// Creates a <see cref="NPOI.XSSF.UserModel.XSSFWorkbook"/> to export Excel File as a <see cref="Stream"/>.
        /// </summary>
        /// <typeparam name="TPoco">Data transfer object which has to be marked as <see cref="IExcelPOCO"/> as a standart.</typeparam>
        /// <param name="workbookSettings"><see cref="WorkbookSettings"/> will be used to store detailed information.</param>
        /// <param name="excelData">List of data transfer objects which will be used to create the excel file.</param>
        /// <returns></returns>
        public static async Task<XLResult> ExportAsync<TPoco>(WorkbookSettings workbookSettings, List<TPoco> excelData) where TPoco : class, IExcelPOCO, new()
        {
            var memory = new MemoryStream();

            using (var fs = new MemoryStream())
            {
                var workbookCreator = new XLWorkbookCreator<TPoco>(workbookSettings, excelData);

                XSSFWorkbook workbook = workbookCreator.CreateWorkbook();

                workbook.Write(fs);

                var byteArray = fs.ToArray();
                await memory.WriteAsync(byteArray);
            }
            memory.Position = 0;

            return new XLResult()
            {
                Success = true,
                Data = memory
            };
        }
    }
}
