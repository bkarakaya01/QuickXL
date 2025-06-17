using Ardalis.GuardClauses;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Builders;

public sealed class ExportBuilder<TDto> where TDto : class, new()
{
    internal List<TDto> Data { get; set; }
    internal ColumnBuilder<TDto> ColumnBuilder { get; set; }
    internal WorkbookSettings WorkbookSettings { get;set;}

    internal ExportBuilder(WorkbookSettings settings)
    {
        Data = [];
        ColumnBuilder = new ColumnBuilder<TDto>(this);
        WorkbookSettings = settings ?? throw new ArgumentNullException(nameof(settings));
    }    

    public ColumnBuilder<TDto> WithData(IEnumerable<TDto> data)
    {
        Guard.Against.Null(data);

        Data.AddRange(data);

        return ColumnBuilder;
    }    
}
