using Ardalis.GuardClauses;
using QuickXL.Core.Factory;
using QuickXL.Core.Models;
using QuickXL.Core.Settings;
using QuickXL.Core.Settings.Columns;

namespace QuickXL.Core.Builders;

public sealed class ColumnBuilder<TDto>
    where TDto : class, new()
{
    internal IList<ColumnBuilderItem<TDto>> ColumnBuilderItems { get; set; }

    internal ExportBuilder<TDto> ExportBuilder;

    internal ColumnBuilder(ExportBuilder<TDto> exportBuilder)
    {
        ExportBuilder = exportBuilder;
        ColumnBuilderItems = [];
    }


    public ColumnBuilder<TDto> AddColumn(string header, Func<TDto, object> propertySelector, Action<ColumnSettings>? configuration = null)
    {
        Guard.Against.Null(ExportBuilder);
        Guard.Against.Null(ColumnBuilderItems);

        ColumnSettings columnSettings = new();

        configuration?.Invoke(columnSettings);

        ColumnBuilderItems.Add(new()
        {
            Header = header,
            PropertySelector = propertySelector,
            ColumnSettings = columnSettings
        });

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