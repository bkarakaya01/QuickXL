using Ardalis.GuardClauses;

namespace QuickXL.Core.Builders;

public sealed class ExportBuilder<TDto> where TDto : class, new()
{
    internal List<TDto> Data { get; set; }
    internal ColumnBuilder<TDto> ColumnBuilder { get; set; }

    internal ExportBuilder()
    {
        Data = [];
        ColumnBuilder = new ColumnBuilder<TDto>(this);
    }

    public ColumnBuilder<TDto> WithData(IEnumerable<TDto> data)
    {
        Guard.Against.Null(data);

        Data.AddRange(data);

        return ColumnBuilder;
    }
}
