using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Ardalis.GuardClauses;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Core.Builders;
using QuickXL.Core.Settings;

namespace QuickXL.Core.Helpers
{
    /// <summary>
    /// Orchestrates the creation of an Excel workbook using helper classes and OpenXmlWriter.
    /// </summary>
    internal sealed class OpenXmlWorkbookHelper<TDto>
        where TDto : class, new()
    {
        /// <summary>
        /// Creates an Excel file based on the ExportBuilder configuration and returns it as a byte array.
        /// </summary>
        public byte[] CreateWorkbook(ExportBuilder<TDto> builder)
        {
            Guard.Against.Null(builder, nameof(builder));
            WorkbookSettings settings = builder.WorkbookSettings
                           ?? throw new ArgumentNullException(nameof(builder.WorkbookSettings));

            using var ms = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true);

            var (wbPart, wsPart) = InitializeWorkbook(doc, settings.SheetName);
            AddStylesheet(wbPart, builder);
            WriteWorksheetData(wsPart, builder, settings.FirstRowIndex);

            wbPart.Workbook.Save();
            return ms.ToArray();
        }

        /// <summary>
        /// Initializes the workbook and worksheet parts and returns them.
        /// </summary>
        private static (WorkbookPart wbPart, WorksheetPart wsPart) InitializeWorkbook(
            SpreadsheetDocument doc,
            string sheetName)
        {
            var wbPart = doc.AddWorkbookPart();
            wbPart.Workbook = new Workbook();
            var sheets = wbPart.Workbook.AppendChild(new Sheets());
            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet();

            sheets.Append(new Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = sheetName
            });

            return (wbPart, wsPart);
        }

        /// <summary>
        /// Adds the stylesheet to the workbook based on the builder settings.
        /// </summary>
        private static void AddStylesheet(
            WorkbookPart wbPart,
            ExportBuilder<TDto> builder)
        {
            var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = StylesheetFactory.Create(builder);
            stylesPart.Stylesheet.Save();
        }

        /// <summary>
        /// Writes column definitions, headers, and data rows into the given worksheet part.
        /// </summary>
        private static void WriteWorksheetData(
            WorksheetPart wsPart,
            ExportBuilder<TDto> builder,
            int firstRowIndex)
        {
            Guard.Against.Null(wsPart, nameof(wsPart));
            Guard.Against.Null(builder, nameof(builder));
            Guard.Against.Null(builder.Data, nameof(builder.Data));

            var columns = builder.ColumnBuilder.ColumnBuilderItems;
            Guard.Against.Null(columns, nameof(columns));

            int count = columns.Count;
            var columnNames = Enumerable.Range(0, count)
                .Select(ExcelReferenceHelper.GetColumnName)
                .ToArray();
            
            var getters = columns.Select(c => c.ValueGetter).ToArray();

            using var writer = OpenXmlWriter.Create(wsPart);
            writer.WriteStartElement(new Worksheet());

            // Write column widths
            writer.WriteStartElement(new Columns());
            for (int i = 0; i < count; i++)
            {
                double width = ColumnWidthCalculator.Calculate(
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

            // Write sheet data (headers and rows)
            writer.WriteStartElement(new SheetData());

            // Header row
            int headerRow = firstRowIndex + 1;
            writer.WriteStartElement(new Row { RowIndex = (uint)headerRow });
            for (int i = 0; i < count; i++)
            {
                writer.WriteElement(new Cell
                {
                    CellReference = ExcelReferenceHelper.GetCellReference(columnNames[i], headerRow),
                    DataType = CellValues.String,
                    StyleIndex = 1U,
                    CellValue = new CellValue(columns[i].HeaderName)
                });
            }
            writer.WriteEndElement();

            // Data rows
            int rowIndex = headerRow + 1;
            foreach (var item in builder.Data)
            {
                writer.WriteStartElement(new Row { RowIndex = (uint)rowIndex });
                for (int i = 0; i < count; i++)
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
    }
}
