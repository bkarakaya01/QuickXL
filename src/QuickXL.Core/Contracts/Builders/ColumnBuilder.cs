using Ardalis.GuardClauses;
using QuickXL.Core.Factory;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Builders;

public sealed class ColumnBuilder<TDto>
    where TDto : class, new()
{
    internal Dictionary<string, Func<TDto, object>> HeaderPropertySelectors { get; set; }

    internal ExportBuilder<TDto> ExportBuilder;

    internal ColumnBuilder(ExportBuilder<TDto> exportBuilder)
    {
        ExportBuilder = exportBuilder;
        HeaderPropertySelectors = [];
    }


    public ColumnBuilder<TDto> AddColumn(string header, Func<TDto, object> propertySelector)
    {
        Guard.Against.Null(ExportBuilder);
        Guard.Against.Null(HeaderPropertySelectors);

        HeaderPropertySelectors.Add(header, propertySelector);

        return this;
    }

    public Exporter<TDto> Build(Action<WorkbookSettings>? configuration = null)
    {
        Guard.Against.Null(ExportBuilder);

        WorkbookSettings workbookSettings = new();

        configuration?.Invoke(workbookSettings);

        return ExporterFactory<TDto>.CreateExporter(ExportBuilder!, workbookSettings);
    }
}
