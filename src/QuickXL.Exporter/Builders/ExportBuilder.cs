using Ardalis.GuardClauses;

namespace QuickXL;

public sealed class ExportBuilder<TDto> where TDto : class, new()
{
    internal Dictionary<string, Func<TDto, object>> HeaderPropertySelectors;
    internal List<TDto> Data;

    public ExportBuilder()
    {
        HeaderPropertySelectors = [];
        Data = [];
    }

    public ExportBuilder<TDto> WithData(IEnumerable<TDto> data)
    {
        Guard.Against.Null(data);
        
        Data.AddRange(data);

        return this;
    }

    public ExportBuilder<TDto> AddColumn(string header, Func<TDto, object> propertySelector)
    {
        Guard.Against.Null(HeaderPropertySelectors);

        HeaderPropertySelectors.Add(header, propertySelector);

        return this;
    }

    public Exporter<TDto> Build(Action<ExportSettings> configuration)
    {
        ExportSettings exportSettings = new();

        configuration?.Invoke(exportSettings);

        return new()
        {
            ExportBuilder = this,
            WorkbookSettings = exportSettings
        };
    }
}
