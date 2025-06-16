using Ardalis.GuardClauses;
using QuickXL.Core.Builders;
using QuickXL.Core.Helpers;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Services;

internal class QuickXLExporter : IXLExporter
{
    private readonly WorkbookSettings _settings;

    public QuickXLExporter(QuickXLOptions options)
    {
        _settings = new WorkbookSettings
        {
            SheetName = options.DefaultSheetName,
            FirstRowIndex = options.DefaultFirstRowIndex
        };
    }

    public byte[] Export<TDto>(Action<ExportBuilder<TDto>> buildAction)
        where TDto : class, new()
    {
        ExportBuilder<TDto> builder = new();
        buildAction(builder);

        OpenXmlWorkbookHelper<TDto> helper = new();
        return helper.CreateWorkbook(builder, _settings);
    }

    public byte[] Export<TDto>(
    IEnumerable<TDto> data,
    Action<ColumnBuilder<TDto>> configureColumns)
    where TDto : class, new()
    {
        Guard.Against.Null(data, nameof(data));
        Guard.Against.Null(configureColumns, nameof(configureColumns));

        ExportBuilder<TDto> exportBuilder = new();
        
        exportBuilder.WithData(data);
        
        configureColumns(exportBuilder.ColumnBuilder);
        
        return new OpenXmlWorkbookHelper<TDto>()
               .CreateWorkbook(exportBuilder, _settings);
    }
}
