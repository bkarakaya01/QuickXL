using Ardalis.GuardClauses;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace QuickXL.Exporter.Helpers
{

    /// <summary>
    /// Builds an Excel workbook (in-memory) from an ExportBuilder, using DOM APIs.
    /// </summary>
    internal sealed class OpenXmlWorkbookHelper<TDto>
            where TDto : class, new()
    {
        public byte[] CreateWorkbook(ExportBuilder<TDto> builder)
        {
            Guard.Against.Null(builder, nameof(builder));
            var settings = builder.WorkbookSettings
                           ?? throw new ArgumentNullException(nameof(builder.WorkbookSettings));

            using var ms = new MemoryStream();

            // NOTE: I guess the using scope is mandatory in here.
            using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true))
            {
                // 1) WorkbookPart + Stylesheet
                var wbPart = document.AddWorkbookPart();
                wbPart.Workbook = new Workbook();

                var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = StylesheetFactory.Create(builder);
                stylesPart.Stylesheet.Save();

                // 2) WorksheetPart + içerik
                var wsPart = wbPart.AddNewPart<WorksheetPart>();
                var worksheet = new Worksheet();
                var sheetData = new SheetData();

                // Columns
                var columns = builder.ColumnBuilder.ColumnBuilderItems;
                int colCount = columns.Count;
                var getters = columns.Select(c => c.ValueGetter).ToArray();
                var colNames = Enumerable.Range(0, colCount)
                                         .Select(ExcelReferenceHelper.GetColumnName)
                                         .ToArray();

                var cols = new Columns();
                for (int i = 0; i < colCount; i++)
                {
                    double w = ColumnWidthCalculator.Calculate(
                        builder.Data, getters[i], columns[i].HeaderName.Length);
                    cols.Append(new Column
                    {
                        Min = (uint)(i + 1),
                        Max = (uint)(i + 1),
                        Width = w,
                        CustomWidth = true
                    });
                }
                worksheet.Append(cols);

                // Header
                int headerIdx = settings.FirstRowIndex;
                var headerRow = new Row { RowIndex = (uint)(headerIdx + 1) };
                for (int i = 0; i < colCount; i++)
                {
                    headerRow.Append(new Cell
                    {
                        CellReference = $"{colNames[i]}{headerIdx + 1}",
                        DataType = CellValues.String,
                        StyleIndex = 1U,
                        CellValue = new CellValue(columns[i].HeaderName)
                    });
                }
                sheetData.Append(headerRow);

                // Data
                int rowIdx = headerIdx + 2;
                foreach (var item in builder.Data)
                {
                    var row = new Row { RowIndex = (uint)rowIdx };
                    for (int i = 0; i < colCount; i++)
                    {
                        var txt = getters[i](item)?.ToString() ?? "";
                        row.Append(new Cell
                        {
                            CellReference = $"{colNames[i]}{rowIdx}",
                            DataType = CellValues.String,
                            StyleIndex = 2U,
                            CellValue = new CellValue(txt)
                        });
                    }
                    sheetData.Append(row);
                    rowIdx++;
                }

                // Attach sheetData & save worksheet
                worksheet.Append(sheetData);
                wsPart.Worksheet = worksheet;
                wsPart.Worksheet.Save();

                // Hook into sheets + save workbook
                var sheets = wbPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new Sheet
                {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1U,
                    Name = settings.SheetName
                });
                wbPart.Workbook.Save();

                // This is where document.Dispose() called.
                // which writes all the XML required for package closure to the MemoryStream.
            }

            // Now the package is completely closed, the ZIP footer is written, our content is ready.
            return ms.ToArray();
        }
    }
}