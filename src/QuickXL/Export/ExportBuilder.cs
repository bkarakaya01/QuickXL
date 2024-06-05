using Ardalis.GuardClauses;
using QuickXL.Core.Contracts;
using QuickXL.Core.Contracts.Builders;
using QuickXL.Core.Contracts.Settings;
using QuickXL.Infrastructure.Export.Factory;
using QuickXL.Infrastructure.Export.Settings;

namespace QuickXL;

public sealed class ExportBuilder<TDto> : IExportBuilder<TDto> 
    where TDto : class, new()
{
    public Dictionary<string, Func<TDto, object>> HeaderPropertySelectors { get; set; }
    public List<TDto> Data { get; set; }

    public ExportBuilder()
    {
        HeaderPropertySelectors = [];
        Data = [];
    }

    public IExportBuilder<TDto> WithData(IEnumerable<TDto> data)
    {
        Guard.Against.Null(data);

        Data.AddRange(data);

        return this;
    }

    public IExportBuilder<TDto> AddColumn(string header, Func<TDto, object> propertySelector)
    {
        Guard.Against.Null(Data);
        Guard.Against.Null(HeaderPropertySelectors);

        HeaderPropertySelectors.Add(header, propertySelector);

        return this;
    }

    public IExporter<TDto> Build(Action<IWorkbookSettings>? configuration = null)
    {
        WorkbookSettings workbookSettings = new();

        configuration?.Invoke(workbookSettings);

        return ExporterFactory<TDto>.CreateExporter(this, workbookSettings);
    }
}
