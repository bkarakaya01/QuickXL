using Ardalis.GuardClauses;
using NPOI.XSSF.UserModel;
using QuickXL.Core.Contracts;
using QuickXL.Core.Contracts.Builders;
using QuickXL.Core.Contracts.Settings;
using QuickXL.Core.Extensions.IO;
using QuickXL.Core.Result;
using QuickXL.Infrastructure.Export.Helpers;

namespace QuickXL.Infrastructure.Export;

internal sealed class Exporter<TDto> : IExporter<TDto>
    where TDto : class, new()
{
    public IExportBuilder<TDto>? ExportBuilder { get; set; }
    
    internal readonly IWorkbookSettings WorkbookSettings;

    internal Exporter(IWorkbookSettings workbookSettings)
    {
        WorkbookSettings = workbookSettings;
    }

    /// <summary>
    /// Creates a <see cref="XSSFWorkbook"/> to export Excel File as a <see cref="Stream"/>.
    /// 
    /// <para>
    ///    Stream will be stored in <see cref="XLResult.Data"/>.
    /// </para>
    /// </summary>
    /// <typeparam name="TDto">Data transfer object</typeparam>
    /// <param name="configuration"><see cref="Settings.WorkbookSettings"/> object will be used to build metada.</param>
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
                XSSFWorkbook workbook = XLWorkbookHelper<TDto>.Instance.CreateWorkbook(ExportBuilder!, WorkbookSettings);

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
