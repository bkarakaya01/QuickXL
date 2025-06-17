using Ardalis.GuardClauses;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Core.Builders;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Helpers
{
    internal sealed class OpenXmlWorkbookHelper<TDto>
        where TDto : class, new()
    {
        public byte[] CreateWorkbook(ExportBuilder<TDto> exportBuilder)
        {
            Guard.Against.Null(exportBuilder, nameof(exportBuilder));
            WorkbookSettings? settings = exportBuilder.WorkbookSettings
                           ?? throw new ArgumentNullException(nameof(exportBuilder.WorkbookSettings));

            var columns = exportBuilder.ColumnBuilder.ColumnBuilderItems;
            int colCount = columns.Count;
            if (colCount == 0)
                throw new InvalidOperationException("En az bir sütun tanımlanmalıdır.");

            // Sütun harflerini önceden hesapla
            var columnRefs = Enumerable.Range(0, colCount)
                                       .Select(ComputeColumnName)
                                       .ToArray();

            // Lambda'ları derle ve sakla
            var getters = columns
                .Select(c => c.ValueGetter)
                .ToArray();

            using var ms = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
            {
                // Workbook oluştur
                var wbPart = document.AddWorkbookPart();
                wbPart.Workbook = new Workbook();

                // Stylesheet ekle
                var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = GenerateStylesheet(exportBuilder);
                stylesPart.Stylesheet.Save();

                // Sheets koleksiyonu
                var sheets = wbPart.Workbook.AppendChild(new Sheets());
                var wsPart = wbPart.AddNewPart<WorksheetPart>();

                // Streaming API ile SheetData oluştur
                using (var writer = OpenXmlWriter.Create(wsPart))
                {
                    writer.WriteStartElement(new Worksheet());

                    // Columns (genişlik)
                    writer.WriteStartElement(new Columns());
                    for (int c = 0; c < colCount; c++)
                    {
                        double width = CalculateApproxWidth(
                            exportBuilder.Data, getters[c], columns[c].HeaderName.Length);
                        writer.WriteElement(new Column
                        {
                            Min = (uint)(c + 1),
                            Max = (uint)(c + 1),
                            Width = width,
                            CustomWidth = true
                        });
                    }
                    writer.WriteEndElement(); // Columns

                    // SheetData başlangıcı
                    writer.WriteStartElement(new SheetData());

                    // Header satırı
                    int headerRow = settings.FirstRowIndex + 1;
                    writer.WriteStartElement(new Row { RowIndex = (uint)headerRow });
                    for (int c = 0; c < colCount; c++)
                    {
                        writer.WriteElement(new Cell
                        {
                            CellReference = $"{columnRefs[c]}{headerRow}",
                            DataType = CellValues.String,
                            StyleIndex = 1U,
                            CellValue = new CellValue(columns[c].HeaderName)
                        });
                    }
                    writer.WriteEndElement(); // Row

                    // Veri satırları
                    int rowIdx = headerRow + 1;
                    foreach (var item in exportBuilder.Data)
                    {
                        writer.WriteStartElement(new Row { RowIndex = (uint)rowIdx });
                        for (int c = 0; c < colCount; c++)
                        {
                            var raw = getters[c](item);
                            var text = raw?.ToString() ?? string.Empty;
                            writer.WriteElement(new Cell
                            {
                                CellReference = $"{columnRefs[c]}{rowIdx}",
                                DataType = CellValues.String,
                                StyleIndex = 2U,
                                CellValue = new CellValue(text)
                            });
                        }
                        writer.WriteEndElement(); // Row
                        rowIdx++;
                    }

                    writer.WriteEndElement(); // SheetData
                    writer.WriteEndElement(); // Worksheet
                }

                // Sheet tanımını ekle
                sheets.Append(new Sheet
                {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1,
                    Name = settings.SheetName
                });

                wbPart.Workbook.Save();
            }

            return ms.ToArray();
        }

        private static Stylesheet GenerateStylesheet(ExportBuilder<TDto> exportBuilder)
        {
            var gs = exportBuilder.ColumnBuilder.XLGeneralStyle;

            var fonts = new Fonts(
                new Font(), // default
                new Font(
                    new Bold(),
                    new FontSize { Val = gs?.HeaderStyle.FontSize ?? 11 },
                    new Color
                    {
                        Rgb = new HexBinaryValue(
                            (gs?.HeaderStyle.ForegroundColor ?? "#000000")
                                .TrimStart('#'))
                    },
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
                new CellFormat(),                            // index=0 default
                new CellFormat { FontId = 1, ApplyFont = true }, // index=1 header
                new CellFormat { FontId = 0, ApplyFont = true }  // index=2 data
            );

            return new Stylesheet(fonts, fills, borders, cellStyleFormats, cellFormats);
        }

        private static double CalculateApproxWidth(
            IEnumerable<TDto> data,
            Func<TDto, object?> getter,
            int headerLength)
        {
            int max = headerLength;
            foreach (var item in data)
            {
                var text = getter(item)?.ToString() ?? string.Empty;
                if (text.Length > max) max = text.Length;
            }
            return Math.Min(max * 1.2, 255);
        }

        private static string ComputeColumnName(int colIndex)
        {
            string name = string.Empty;
            int dividend = colIndex + 1;
            while (dividend > 0)
            {
                int mod = (dividend - 1) % 26;
                name = (char)('A' + mod) + name;
                dividend = (dividend - mod) / 26;
            }
            return name;
        }
    }
}
