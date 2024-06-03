using Ardalis.GuardClauses;
using NPOI.XSSF.UserModel;
using QuickXL.Core.Extensions.IO;
using QuickXL.Core.Result;
using System.IO;

namespace QuickXL;

public sealed class Exporter<TDto> where TDto : class, new()
{
    internal ExportBuilder<TDto>? ExportBuilder { get; set; }
    internal ExportSettings WorkbookSettings { get; set; }
    internal Exporter()
    {
        WorkbookSettings = new();
    }

    /// <summary>
    /// Creates a <see cref="XSSFWorkbook"/> to export Excel File as a <see cref="Stream"/>.
    /// 
    /// <para>
    ///    Stream will be stored in <see cref="XLResult.Data"/>.
    /// </para>
    /// </summary>
    /// <typeparam name="TDto">Data transfer object</typeparam>
    /// <param name="configuration"><see cref="ExportSettings"/> object will be used to build metada.</param>
    /// <param name="excelData">List data to export.</param>
    /// <returns></returns>
    public XLResult Export()
    {
        try
        {            
            ValidateOrThrow();

            using FileStream fs = new(Path.Combine(@"C:\Users\bkara\Projects\", "QuickXL.xlsx"), FileMode.Create, FileAccess.ReadWrite);

            XLWorkbookCreator<TDto> workbookCreator = new(this);

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

    private void ValidateOrThrow()
    {
        Guard.Against.Null(ExportBuilder);
        Guard.Against.Null(WorkbookSettings);
    }
}
