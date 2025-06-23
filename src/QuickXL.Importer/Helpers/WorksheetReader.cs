using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace QuickXL.Importer.Helpers
{
    /// <summary>
    /// Provides methods to read worksheet rows and cell values from an OpenXML spreadsheet.
    /// </summary>
    internal static class WorksheetReader
    {
        /// <summary>
        /// Opens the given Excel stream and retrieves all rows from the first worksheet's SheetData.
        /// </summary>
        /// <param name="source">A <see cref="Stream"/> containing the .xlsx data.</param>
        /// <param name="wbPart">
        /// Upon return, contains the <see cref="WorkbookPart"/> associated with the opened document.
        /// </param>
        /// <returns>
        /// A list of <see cref="Row"/> elements representing each row in the worksheet.
        /// </returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="source"/> is null.</exception>
        internal static List<Row> ReadRows(Stream source, out WorkbookPart wbPart)
        {
            if (source == null)
                throw new ArgumentNullException(nameof(source));

            var document = SpreadsheetDocument.Open(source, false);
            wbPart = document.WorkbookPart!;

            // Currently selects the first sheet; can be extended to use sheet settings
            var sheet = wbPart.Workbook.Descendants<Sheet>().First();
            var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!);

            return wsPart.Worksheet
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .ToList();
        }

        /// <summary>
        /// Retrieves the text content of a cell, correctly handling shared string lookups.
        /// </summary>
        /// <param name="cell">The <see cref="Cell"/> to read.</param>
        /// <param name="wbPart">The <see cref="WorkbookPart"/> for resolving shared strings.</param>
        /// <returns>The cell's string value.</returns>
        /// <exception cref="ArgumentNullException">
        /// Thrown if <paramref name="cell"/> or <paramref name="wbPart"/> is null.
        /// </exception>
        internal static string GetCellValue(Cell cell, WorkbookPart wbPart)
        {
            if (cell == null)
                throw new ArgumentNullException(nameof(cell));
            if (wbPart == null)
                throw new ArgumentNullException(nameof(wbPart));

            var raw = cell.CellValue?.InnerText ?? string.Empty;
            if (cell.DataType?.Value == CellValues.SharedString
                && wbPart.SharedStringTablePart != null)
            {
                var sst = wbPart.SharedStringTablePart.SharedStringTable;
                return sst.ElementAt(int.Parse(raw)).InnerText;
            }
            return raw;
        }
    }
}
