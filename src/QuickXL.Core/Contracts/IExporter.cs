using QuickXL.Core.Contracts.Builders;
using QuickXL.Core.Result;

namespace QuickXL.Core.Contracts
{
    public interface IExporter<TDto> where TDto : class, new()
    {
        IExportBuilder<TDto>? ExportBuilder { get; set; }
        XLResult Export();
    }
}
