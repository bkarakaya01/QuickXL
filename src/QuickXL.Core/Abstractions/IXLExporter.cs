using QuickXL.Core.Builders;

namespace QuickXL;

public interface IXLExporter
{
    byte[] Export<TDto>(Action<ExportBuilder<TDto>> buildAction)
        where TDto : class, new();

    byte[] Export<TDto>(
            IEnumerable<TDto> data,
            Action<ColumnBuilder<TDto>> configureColumns)
            where TDto : class, new();
}
