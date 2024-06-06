using QuickXL.Core.Builders;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Factory;

public static class ExporterFactory<TDto> where TDto : class, new()
{
    public static ExportBuilder<TDto> CreateBuilder() => new();

    public static Exporter<TDto> CreateExporter(ExportBuilder<TDto> exportBuilder, WorkbookSettings workbookSettings) => 
        new(exportBuilder, workbookSettings);
}
