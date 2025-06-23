namespace QuickXL.Exporter
{
    public sealed class ExportBuilder<TDto>
        where TDto : class, new()
    {
        internal List<TDto> Data { get; set; }
        internal ColumnBuilder<TDto> ColumnBuilder { get; set; }
        
        public WorkbookSettings WorkbookSettings { get; set; }

        public ExportBuilder(WorkbookSettings settings)
        {
            Data = [];
            ColumnBuilder = new ColumnBuilder<TDto>(this);
            WorkbookSettings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        public ColumnBuilder<TDto> WithData(IEnumerable<TDto> data)
        {
            if (data == null)
                throw new ArgumentNullException(nameof(data));

            Data.AddRange(data);

            return ColumnBuilder;
        }
    }
}
