using QuickXL.Core.Contracts;
using QuickXL.Core.Contracts.Builders;
using QuickXL.Core.Contracts.Settings;

namespace QuickXL.Infrastructure.Export.Factory;

public static class ExporterFactory<TDto> where TDto : class, new()
{
    public static IExporter<TDto> CreateExporter(IExportBuilder<TDto> exportBuilder, IWorkbookSettings exportSettings)
    {
        return new Exporter<TDto>(exportSettings)
        {
            ExportBuilder = exportBuilder
        };
    }
}
