using NPOI.XSSF.UserModel;
using QuickXL.Core.Extensions.IO;
using QuickXL.Core.Result;

namespace QuickXL;

public sealed class Exporter
{
    /// <summary>
    /// Creates a <see cref="XSSFWorkbook"/> to export Excel File as a <see cref="Stream"/>.
    /// 
    /// <para>
    ///    Stream will be stored in <see cref="XLResult.Data"/>.
    /// </para>
    /// </summary>
    /// <typeparam name="TDto">Data transfer object</typeparam>
    /// <param name="exportSettings"><see cref="ExportSettings"/> object will be used to build metada.</param>
    /// <param name="excelData">List data to export.</param>
    /// <returns></returns>
    public static XLResult Export<TDto>(ExportSettings exportSettings, IList<TDto> excelData) where TDto : class, new()
    {
        try
        {
            using var fs = new MemoryStream();

            var workbookCreator = new XLWorkbookCreator<TDto>(exportSettings, excelData);

            XSSFWorkbook workbook = workbookCreator.CreateWorkbook();

            workbook.Write(fs);

            fs.Reset();

            return XLResult.Success(fs);
        }
        catch (Exception ex)
        {
            return (XLResult)ex;
        }
    }        
}
