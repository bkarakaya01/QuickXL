namespace QuickXL.Importer
{
    public class XLImporter
    {
        /// <summary>
        /// Start import from a Stream.
        /// </summary>
        public static ImportBuilder<TDto> Create<TDto>(Action<ImportSettings>? configuration = null)
            where TDto : class, new()
        {
            ImportSettings settings = new();

            configuration?.Invoke(settings);

            return new ImportBuilder<TDto>(settings);
        }
    }
}