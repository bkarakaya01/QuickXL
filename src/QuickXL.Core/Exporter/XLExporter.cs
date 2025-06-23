namespace QuickXL.Exporter
{
    public class XLExporter
    {
        public static ExportBuilder<TDto> Create<TDto>(Action<WorkbookSettings>? configuration = null)
            where TDto : class, new()
        {
            WorkbookSettings settings = new();

            configuration?.Invoke(settings);

            return new ExportBuilder<TDto>(settings);
        }
    }
}
