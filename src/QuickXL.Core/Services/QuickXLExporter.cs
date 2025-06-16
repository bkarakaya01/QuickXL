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
        // 1) Builder oluşturup data’yı set et
        var builder = new ExportBuilder<TDto>()
            .WithData(data);

        // 2) Sütun konfigürasyonunu uygula
        configureColumns(builder.ColumnBuilder);

        // 3) Dosyayı oluştur
        return new XLWorkbookHelper<TDto>()
               .CreateWorkbook(builder, _defaults);
    }
}
}
