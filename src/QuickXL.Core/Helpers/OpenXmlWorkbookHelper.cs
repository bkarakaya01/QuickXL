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

namespace QuickXL.Core.Helpers;

/// <summary>
/// Orchestrates the creation of an Excel workbook using OpenXmlWriter.
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
        var settings = builder.WorkbookSettings
                       ?? throw new ArgumentNullException(nameof(builder.WorkbookSettings));

        using var ms = new MemoryStream();
        using var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, true);

        // Initialize workbook and worksheet parts
        var (wbPart, wsPart) = InitializeWorkbook(doc, settings.SheetName);

        // Add styles
        AddStylesheet(wbPart, builder);

        // Write content (columns, headers, data)
        WriteWorksheetData(wsPart, builder, settings.FirstRowIndex);

        // Save the workbook XML
        wbPart.Workbook.Save();
        return ms.ToArray();
    }

    /// <summary>
    /// Adds a WorkbookPart and an empty WorksheetPart, returning both.
    /// </summary>
    private static (WorkbookPart wbPart, WorksheetPart wsPart) InitializeWorkbook(
        SpreadsheetDocument document,
        string sheetName)
    {
        // 1) Create the workbook part
        var wbPart = document.AddWorkbookPart();
        wbPart.Workbook = new Workbook();

        // 2) Create the sheets container
        var sheets = wbPart.Workbook.AppendChild(new Sheets());

        // 3) Create a worksheet part WITHOUT pre-populating it
        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        // Do NOT set wsPart.Worksheet here when using OpenXmlWriter!

        // 4) Link the worksheet into the workbook
        var sheetId = (uint)1;
        sheets.Append(new Sheet
        {
            Id = wbPart.GetIdOfPart(wsPart),
            SheetId = sheetId,
            Name = sheetName
        });

        return (wbPart, wsPart);
    }

    /// <summary>
    /// Inserts the stylesheet into the workbook based on builder settings.
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
    /// Streams column definitions, header row and data rows into the given worksheet part.
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
        Guard.Against.NullOrEmpty(columns, nameof(columns));

        int count = columns.Count;
        var colNames = Enumerable.Range(0, count)
                                    .Select(ExcelReferenceHelper.GetColumnName)
                                    .ToArray();
        var getters = columns.Select(c => c.ValueGetter).ToArray();

        using var writer = OpenXmlWriter.Create(wsPart);
        writer.WriteStartElement(new Worksheet());

        // 1) Columns (width definitions)
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
        writer.WriteEndElement(); // Columns

        // 2) SheetData (header + data rows)
        writer.WriteStartElement(new SheetData());

        // 2a) Header row
        int headerRow = firstRowIndex + 1;
        writer.WriteStartElement(new Row { RowIndex = (uint)headerRow });
        for (int i = 0; i < count; i++)
        {
            writer.WriteElement(new Cell
            {
                CellReference = ExcelReferenceHelper.GetCellReference(colNames[i], headerRow),
                DataType = CellValues.String,
                StyleIndex = 1U,
                CellValue = new CellValue(columns[i].HeaderName)
            });
        }
        writer.WriteEndElement(); // Row

        // 2b) Data rows
        int rowIndex = headerRow + 1;
        foreach (var item in builder.Data)
        {
            writer.WriteStartElement(new Row { RowIndex = (uint)rowIndex });
            for (int i = 0; i < count; i++)
            {
                var text = getters[i](item)?.ToString() ?? string.Empty;
                writer.WriteElement(new Cell
                {
                    CellReference = ExcelReferenceHelper.GetCellReference(colNames[i], rowIndex),
                    DataType = CellValues.String,
                    StyleIndex = 2U,
                    CellValue = new CellValue(text)
                });
            }
            writer.WriteEndElement(); // Row
            rowIndex++;
        }

        writer.WriteEndElement(); // SheetData
        writer.WriteEndElement(); // Worksheet
    }
}