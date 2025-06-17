using Ardalis.GuardClauses;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Core.Builders;

namespace QuickXL.Core.Helpers;

/// <summary>
/// Orchestrates creation of an Excel workbook using helper classes and OpenXmlWriter.
/// </summary>
internal sealed class OpenXmlWorkbookHelper<TDto>
    where TDto : class, new()
{
    /// <summary>
    /// Builds the Excel workbook as a byte array.
    /// </summary>
    public byte[] CreateWorkbook(ExportBuilder<TDto> builder)
    {
        Guard.Against.Null(builder, nameof(builder));
        var settings = builder.WorkbookSettings ?? throw new ArgumentNullException(nameof(builder.WorkbookSettings));

        var columns = builder.ColumnBuilder.ColumnBuilderItems;
        Guard.Against.NullOrEmpty(columns, nameof(columns));

        var columnCount = columns.Count;
        var columnNames = Enumerable.Range(0, columnCount)
                                     .Select(ExcelReferenceHelper.GetColumnName)
                                     .ToArray();
        var getters = columns.Select(c => c.ValueGetter).ToArray();

        using var ms = new MemoryStream();
        using var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true);

        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();

        // add stylesheet
        var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = StylesheetFactory.Create(builder);
        stylesPart.Stylesheet.Save();

        var sheets = wbPart.Workbook.AppendChild(new Sheets());
        var wsPart = wbPart.AddNewPart<WorksheetPart>();

        using (var writer = OpenXmlWriter.Create(wsPart))
        {
            writer.WriteStartElement(new Worksheet());
            // columns
            writer.WriteStartElement(new Columns());
            for (int i = 0; i < columnCount; i++)
            {
                var width = ColumnWidthCalculator.Calculate(
                    builder.Data,
                    getters[i],
                    columns[i].HeaderName.Length);
                writer.WriteElement(new Column
                {
                    Min = (uint)(i + 1),
                    Max = (uint)(i + 1),
                    Width = width,
                    CustomWidth = true
                });
            }
            writer.WriteEndElement();

            // sheet data
            writer.WriteStartElement(new SheetData());
            // header row
            int headerRowIdx = settings.FirstRowIndex + 1;
            writer.WriteStartElement(new Row { RowIndex = (uint)headerRowIdx });
            for (int i = 0; i < columnCount; i++)
            {
                writer.WriteElement(new Cell
                {
                    CellReference = ExcelReferenceHelper.GetCellReference(columnNames[i], headerRowIdx),
                    DataType = CellValues.String,
                    StyleIndex = 1U,
                    CellValue = new CellValue(columns[i].HeaderName)
                });
            }
            writer.WriteEndElement();

            // data rows
            int rowIndex = headerRowIdx + 1;
            foreach (var item in builder.Data)
            {
                writer.WriteStartElement(new Row { RowIndex = (uint)rowIndex });
                for (int i = 0; i < columnCount; i++)
                {
                    var text = getters[i](item)?.ToString() ?? string.Empty;
                    writer.WriteElement(new Cell
                    {
                        CellReference = ExcelReferenceHelper.GetCellReference(columnNames[i], rowIndex),
                        DataType = CellValues.String,
                        StyleIndex = 2U,
                        CellValue = new CellValue(text)
                    });
                }
                writer.WriteEndElement();
                rowIndex++;
            }

            writer.WriteEndElement(); // SheetData
            writer.WriteEndElement(); // Worksheet
        }

        sheets.Append(new Sheet
        {
            Id = wbPart.GetIdOfPart(wsPart),
            SheetId = 1,
            Name = settings.SheetName
        });

        wbPart.Workbook.Save();
        return ms.ToArray();
    }
}
