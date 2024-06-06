using QuickXL.Core.Factory;
using QuickXL.Infrastructure.Export.Builders;

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
