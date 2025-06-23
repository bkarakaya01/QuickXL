using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Importer.Helpers;
using System.Linq.Expressions;
using System.Reflection;

namespace QuickXL.Importer
{
    /// <summary>
    /// Fluent builder that configures and executes the import of rows from an Excel worksheet
    /// into a list of <typeparamref name="TDto"/> instances.
    /// </summary>
    /// <typeparam name="TDto">The POCO type to which each row will be mapped.</typeparam>
    public sealed class ImportBuilder<TDto>
        where TDto : class, new()
    {
        private readonly ImportSettings _settings;

        /// <summary>
        /// The input stream containing the Excel document. Either this or a file path
        /// must be provided before calling <see cref="Import"/>.
        /// </summary>
        internal Stream? Stream { get; private set; }

        /// <summary>
        /// Initializes a new instance of <see cref="ImportBuilder{TDto}"/>.
        /// </summary>
        /// <param name="settings">Global settings that control sheet selection, header detection, and mapping behavior.</param>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="settings"/> is null.</exception>
        public ImportBuilder(ImportSettings settings)
        {
            _settings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>
        /// Specifies the path of the Excel file to import.
        /// </summary>
        /// <param name="path">The file system path of the `.xlsx` file.</param>
        /// <returns>The same <see cref="ImportBuilder{TDto}"/> instance for chaining.</returns>
        public ImportBuilder<TDto> FromFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("File path must not be empty", nameof(path));

            Stream = File.OpenRead(path);
            return this;
        }

        /// <summary>
        /// Specifies an existing <see cref="Stream"/> that contains the Excel document.
        /// </summary>
        /// <param name="stream">The stream to read the `.xlsx` data from.</param>
        /// <returns>The same <see cref="ImportBuilder{TDto}"/> instance for chaining.</returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="stream"/> is null.</exception>
        public ImportBuilder<TDto> FromStream(Stream stream)
        {
            Stream = stream ?? throw new ArgumentNullException(nameof(stream));
            return this;
        }

        /// <summary>
        /// Manually maps a DTO property to a specific zero-based Excel column index.
        /// </summary>
        /// <typeparam name="TProp">The type of the property being mapped.</typeparam>
        /// <param name="propertySelector">
        /// An expression selecting the target property on <typeparamref name="TDto"/> 
        /// (e.g. <c>x => x.MyProperty</c>).
        /// </param>
        /// <param name="columnIndex">
        /// The zero-based index of the column in the worksheet to bind this property to.
        /// </param>
        /// <returns>The same <see cref="ImportBuilder{TDto}"/> instance for fluent chaining.</returns>
        /// <exception cref="ArgumentNullException">
        /// Thrown if <paramref name="propertySelector"/> is <c>null</c>.
        /// </exception>
        /// <exception cref="ArgumentOutOfRangeException">
        /// Thrown if <paramref name="columnIndex"/> is negative.
        /// </exception>
        /// <exception cref="ArgumentException">
        /// Thrown if the <paramref name="propertySelector"/> body is not a valid member-access
        /// expression or does not refer to a property.
        /// </exception>
        public ImportBuilder<TDto> Map<TProp>(
            Expression<Func<TDto, TProp>> propertySelector,
            int columnIndex)
        {
            if (propertySelector is null)
                throw new ArgumentNullException(nameof(propertySelector));
            if (columnIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(columnIndex), "Column index must be zero or positive.");

            // Extract the PropertyInfo from the expression tree
            var member = propertySelector.Body as MemberExpression
                ?? throw new ArgumentException("The selector must be a simple member access expression.", nameof(propertySelector));

            var pi = member.Member as PropertyInfo
                ?? throw new ArgumentException("The selector must refer to a property.", nameof(propertySelector));

            // Build the setter delegate and register the mapping
            var setter = MappingHelper.BuildSetter(pi);
            _settings.ManualMappings.Add(new Mapping
            {
                Index = columnIndex,
                Setter = setter
            });

            return this;
        }

        /// <summary>
        /// Performs the import operation: reads the worksheet, detects columns, applies both attribute-based
        /// and manual mappings, and returns a list of populated <typeparamref name="TDto"/> instances.
        /// </summary>
        /// <returns>A <see cref="List{TDto}"/> containing one instance per non-header row.</returns>
        /// <exception cref="InvalidOperationException">
        /// Thrown if neither <see cref="FromFile"/> nor <see cref="FromStream"/> was called first.
        /// </exception>
        public List<TDto> Import()
        {
            if (Stream == null)
                throw new InvalidOperationException("No input source specified. Call FromFile() or FromStream() before Import().");

            // 1) Read all rows from the worksheet
            var rows = WorksheetReader.ReadRows(Stream, out var wbPart);

            // 2) Determine header row index
            int headerIdx = HeaderRowDetector.Detect<TDto>(
                rows,
                _settings.HeaderRowSettings,
                rows[_settings.HeaderRowSettings?.StartsAt ?? 0]
                    .Elements<Cell>()
                    .Select(c => WorksheetReader.GetCellValue(c, wbPart)));

            // 3) Build mapping list
            var mappings = new List<Mapping>();
            if (_settings.UseAttributes)
                mappings.AddRange(MappingHelper.BuildAttributeMappings<TDto>(
                    rows[headerIdx]
                        .Elements<Cell>()
                        .Select(c => WorksheetReader.GetCellValue(c, wbPart))
                        .ToList()));
            mappings.AddRange(_settings.ManualMappings);

            // 4) Iterate over data rows and populate DTOs
            var result = new List<TDto>();
            foreach (var row in rows.Skip(headerIdx + 1))
            {
                var dto = new TDto();
                var cells = row.Elements<Cell>().ToArray();
                foreach (var m in mappings)
                {
                    if (m.Index < 0 || m.Index >= cells.Length)
                        continue;

                    string text = WorksheetReader.GetCellValue(cells[m.Index], wbPart);
                    m.Setter(dto, text);
                }
                result.Add(dto);
            }

            return result;
        }
    }
}
