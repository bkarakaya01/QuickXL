using Ardalis.GuardClauses;
using QuickXL.Core.Builders;
using QuickXL.Core.Extensions.IO;
using QuickXL.Core.Helpers;
using QuickXL.Core.Result;
using QuickXL.Core.Settings;

namespace QuickXL.Core;

public sealed class Exporter<TDto>
    where TDto : class, new()
{
    internal ExportBuilder<TDto> ExportBuilder;
    
    internal readonly WorkbookSettings WorkbookSettings;

    internal Exporter(ExportBuilder<TDto> exportBuilder, WorkbookSettings workbookSettings)
    {
        ExportBuilder = exportBuilder;
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
                //XSSFWorkbook workbook = XLWorkbookHelper<TDto>.Instance.CreateWorkbook(ExportBuilder!, WorkbookSettings);

                //workbook.Write(fs);

                //var bytes = fs.ToArray();
                //ms.Write(bytes);
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
