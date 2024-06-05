using QuickXL.Core.Contracts.Settings;

namespace QuickXL.Core.Contracts.Builders
{
    public interface IExportBuilder<TDto> where TDto : class, new()
    {
        IExportBuilder<TDto> WithData(IEnumerable<TDto> data);
        IExportBuilder<TDto> AddColumn(string header, Func<TDto, object> propertySelector);
        IExporter<TDto> Build(Action<IWorkbookSettings>? configuration = null);

        Dictionary<string, Func<TDto, object>> HeaderPropertySelectors { get; set; }
        List<TDto> Data { get; set; }
    }
}
