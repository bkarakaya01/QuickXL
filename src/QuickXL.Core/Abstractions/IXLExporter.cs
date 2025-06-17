using QuickXL.Core.Builders;

namespace QuickXL;

public interface IXLExporter
{
    /// <summary>
    /// Yeni bir ColumnBuilder başlatır. 
    /// Zincir sonrası .Build() ile byte[] elde edilir.
    /// </summary>
    ExportBuilder<TDto> CreateBuilder<TDto>() where TDto : class, new();
}
