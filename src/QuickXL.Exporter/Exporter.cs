using Ardalis.GuardClauses;
using NPOI.XSSF.UserModel;
using QuickXL.Core.Contracts;
using QuickXL.Core.Extensions.IO;
using QuickXL.Core.Result;

namespace QuickXL;

internal sealed class Exporter<TDto> : IExporter<TDto> 
    where TDto : class, new()
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

            MemoryStream ms = new();

            using (var fs = new MemoryStream())
            {
                XLWorkbookCreator<TDto> workbookCreator = new(this);

                XSSFWorkbook workbook = workbookCreator.CreateWorkbook();

                workbook.Write(fs);

                var bytes = fs.ToArray();
                ms.Write(bytes);
            };           

            ms.Reset();

            return XLResult.Success(ms);
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
