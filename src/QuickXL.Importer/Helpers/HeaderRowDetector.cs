using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickXL.Importer.Helpers
{
    /// <summary>
    /// Detects the header row index in a worksheet by matching expected column names.
    /// </summary>
    internal static class HeaderRowDetector
    {
        /// <summary>
        /// Determines which row contains the header, based on a manual setting or by
        /// scoring the first few rows against expected header names.
        /// </summary>
        /// <typeparam name="TDto">The target DTO type (unused in detection logic).</typeparam>
        /// <param name="rows">All worksheet rows read from SheetData.</param>
        /// <param name="settings">Optional settings containing an explicit header row index.</param>
        /// <param name="expectedHeaders">A collection of expected header text values.</param>
        /// <returns>
        /// The zero-based index of the header row. If <paramref name="settings"/> specifies
        /// <see cref="Settings.HeaderRowSettings.StartsAt"/>, that value is returned;
        /// otherwise, the row with the highest match count within the first 10 rows is chosen.
        /// </returns>
        internal static int Detect<TDto>(
            List<Row> rows,
            HeaderRowSettings? settings,
            IEnumerable<string> expectedHeaders)
            where TDto : class, new()
        {
            // If a manual header start index is provided, use that first.
            if (settings?.StartsAt is int idx)
                return idx;

            const int maxScan = 10;
            int bestIdx = 0;
            int bestScore = -1;
            int limit = Math.Min(rows.Count, maxScan);

            // Score each row by how many cell values match expected headers.
            for (int i = 0; i < limit; i++)
            {
                var texts = rows[i].Elements<Cell>()
                    .Select(c => c.CellValue?.InnerText ?? string.Empty);

                int score = texts.Count(text =>
                    expectedHeaders.Any(exp =>
                        string.Equals(exp, text, StringComparison.OrdinalIgnoreCase)));

                if (score > bestScore)
                {
                    bestScore = score;
                    bestIdx = i;
                }
            }

            return bestIdx;
        }
    }
}
