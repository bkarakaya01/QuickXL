using QuickXL.Core.Builders;
using QuickXL.Core.Factory;

namespace QuickXL
{
    public class XLExport<TDto> where TDto : class, new()
    {
        public XLExport() { }

        public ExportBuilder<TDto> CreateBuilder()
        {
            return ExporterFactory<TDto>.CreateBuilder();
        }
    }
}
