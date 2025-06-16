using Ardalis.GuardClauses;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Core.Builders;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Helpers
{
    //OpenXmlWorkbookHelper
    internal sealed class OpenXmlWorkbookHelper<TDto>
        where TDto : class, new()
    {
        public byte[] CreateWorkbook(ExportBuilder<TDto> exportBuilder, WorkbookSettings settings)
        {
            Guard.Against.Null(exportBuilder, nameof(exportBuilder));
            Guard.Against.Null(settings, nameof(settings));

            var headers = exportBuilder.ColumnBuilder.ColumnBuilderItems
                                .Select(x => x.HeaderName)
                                .ToList();
            var propNames = exportBuilder.ColumnBuilder.ColumnBuilderItems
                                .Select(x => x.GetPropertyName())
                                .ToList();

            using var ms = new MemoryStream();
            using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
            {
                // 1. Workbook / Worksheet oluştur
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());
                sheets.Append(new Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = settings.SheetName
                });

                // 2. Stylesheet ekle
                var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = GenerateStylesheet(exportBuilder);
                stylesPart.Stylesheet.Save();

                var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // 3. Header satırı
                int headerRowIndex = settings.FirstRowIndex + 1;
                var headerRow = new Row { RowIndex = (uint)headerRowIndex };
                for (int colIdx = 0; colIdx < headers.Count; colIdx++)
                {
                    var cell = CreateTextCell(colIdx, headerRowIndex - 1, headers[colIdx]);
                    cell.StyleIndex = 1; // header format
                    headerRow.Append(cell);
                }
                sheetData!.Append(headerRow);

                // 4. Veri satırları
                int rowIndex = headerRowIndex + 1;
                foreach (var item in exportBuilder.Data)
                {
                    var dataRow = new Row { RowIndex = (uint)rowIndex };
                    for (int colIdx = 0; colIdx < propNames.Count; colIdx++)
                    {
                        var text = GetPropertyValue(item, propNames[colIdx]);
                        var cell = CreateTextCell(colIdx, rowIndex - 1, text);
                        cell.StyleIndex = 2; // data format
                        dataRow.Append(cell);
                    }
                    sheetData.Append(dataRow);
                    rowIndex++;
                }

                // 5. Sütun genişliklerini yaklaşık auto–size yap
                SetColumnWidths(worksheetPart, headers, exportBuilder, settings);

                workbookPart.Workbook.Save();
            }

            return ms.ToArray();
        }

        private static Stylesheet GenerateStylesheet(ExportBuilder<TDto> exportBuilder)
        {
            var gs = exportBuilder.ColumnBuilder.XLGeneralStyle;

            var fonts = new Fonts(
                new Font(), // default font
                new Font(   // header font
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
                new CellFormat(),                         // index=0 default
                new CellFormat { FontId = 1, ApplyFont = true }, // index=1 header
                new CellFormat { FontId = 0, ApplyFont = true }  // index=2 data
            );

            return new Stylesheet(fonts, fills, borders, cellStyleFormats, cellFormats);
        }

        private static Cell CreateTextCell(int colIndex, int rowIndex, string text)
            => new Cell
            {
                CellReference = GetCellReference(colIndex, rowIndex),
                DataType = CellValues.String,
                CellValue = new CellValue(text)
            };

        private static void SetColumnWidths(
            WorksheetPart wsPart,
            List<string> headers,
            ExportBuilder<TDto> exportBuilder,
            WorkbookSettings settings)
        {
            ArgumentNullException.ThrowIfNull(settings);

            int count = headers.Count;
            var maxLens = new int[count];
            for (int i = 0; i < count; i++)
                maxLens[i] = headers[i].Length;

            foreach (var item in exportBuilder.Data)
            {
                for (int i = 0; i < count; i++)
                {
                    var val = GetPropertyValue(
                        item,
                        exportBuilder.ColumnBuilder.ColumnBuilderItems[i]
                            .GetPropertyName());
                    if (val.Length > maxLens[i])
                        maxLens[i] = val.Length;
                }
            }

            var cols = new Columns();
            for (int i = 0; i < count; i++)
            {
                double width = Math.Min(maxLens[i] * 1.2, 255);
                cols.Append(new Column
                {
                    Min = (uint)(i + 1),
                    Max = (uint)(i + 1),
                    Width = width,
                    CustomWidth = true
                });
            }

            wsPart.Worksheet.InsertAt(cols, 0);
        }

        private static string GetCellReference(int colIndex, int rowIndex)
        {
            string colName = "";
            int dividend = colIndex + 1;
            while (dividend > 0)
            {
                int mod = (dividend - 1) % 26;
                colName = Convert.ToChar('A' + mod) + colName;
                dividend = (dividend - mod) / 26;
            }
            return $"{colName}{rowIndex + 1}";
        }

        private static string GetPropertyValue(TDto item, string propertyName)
            => item.GetType()
                   .GetProperty(propertyName)!
                   .GetValue(item)?
                   .ToString() ?? string.Empty;
    }
}
