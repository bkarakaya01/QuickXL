using QuickXL.Core.Result;

namespace QuickXL.Core.Contracts
{
    public interface IExporter<TDto> where TDto : class, new()
    {
        XLResult Export();
    }
}
