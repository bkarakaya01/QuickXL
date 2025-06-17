using Ardalis.GuardClauses;
using QuickXL.Core.Builders;
using QuickXL.Core.Helpers;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Services;

internal class QuickXLExporter : IXLExporter
{
    private readonly WorkbookSettings _settings;

    public QuickXLExporter(QuickXLOptions opts)
    {
        _settings = new WorkbookSettings
        {
            SheetName = opts.DefaultSheetName,
            FirstRowIndex = opts.DefaultFirstRowIndex
        };
    }

    public ExportBuilder<TDto> CreateBuilder<TDto>()
        where TDto : class, new()
    {
        return new ExportBuilder<TDto>(_settings);
    }
}
