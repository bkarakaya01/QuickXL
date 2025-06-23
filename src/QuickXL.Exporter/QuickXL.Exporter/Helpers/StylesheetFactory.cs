using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Exporter;

namespace QuickXL.Exporter.Helpers
{

    /// <summary>
    /// Factory for creating the OpenXML Stylesheet for the workbook.
    /// </summary>
    internal static class StylesheetFactory
    {
        /// <summary>
        /// Generates a Stylesheet with default and header/data formats.
        /// </summary>
        public static Stylesheet Create<TDto>(ExportBuilder<TDto> builder)
            where TDto : class, new()
        {
            var gs = builder.ColumnBuilder.XLGeneralStyle;
            var fonts = new Fonts(
                new Font(), // default
                new Font(
                    new Bold(),
                    new FontSize { Val = gs?.HeaderStyle.FontSize ?? 11 },
                    new Color { Rgb = new HexBinaryValue((gs?.HeaderStyle.ForegroundColor ?? "#000000").TrimStart('#')) },
                    new FontName { Val = gs?.HeaderStyle.FontName ?? "Calibri" }
                )
            );
            var fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            );
            var borders = new Borders(new Border());
            var cellStyleFormats = new CellStyleFormats(new CellFormat());
            var cellFormats = new CellFormats(
                new CellFormat(),                    // default
                new CellFormat { FontId = 1, ApplyFont = true }, // header
                new CellFormat { FontId = 0, ApplyFont = true }  // data
            );
            return new Stylesheet(fonts, fills, borders, cellStyleFormats, cellFormats);
        }
    }
}