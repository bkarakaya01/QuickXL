using Ardalis.GuardClauses;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using QuickXL.Core.Models.Cells;
using QuickXL.Core.Models;
using QuickXL.Core.Models.Colors;

namespace QuickXL.Core.Styles.Columns;

public sealed class XLCellStyle
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Strikeout { get; set; }
    public string FontName { get; set; }
    public short FontSize { get; set; }
    public XLColor? ForegroundColor { get; set; }
    public XLColor? BackgroundColor { get; set; }
    public FillPattern? FillPattern { get; set; }
    public HorizontalAlignment? HorizontalAlignment { get; set; }
    public VerticalAlignment? VerticalAlignment { get; set; }

    public XLCellStyle()
    {
        Bold = false;
        Italic = false;
        Strikeout = false;
        FontName = "Calibri";
        FontSize = 12;
    }

    internal void Apply<TDto>(XLSheet<TDto> xlSheet, XLCell xlCell) where TDto : class, new()
    {
        ICell? cell = xlCell.Cell;

        Guard.Against.Null(xlSheet);
        Guard.Against.Null(cell);

        if (xlSheet.Sheet.Workbook.CreateCellStyle() is not XSSFCellStyle cellStyle)
            throw new NullReferenceException(nameof(cellStyle));

        // Font settings
        var font = xlSheet.Sheet.Workbook.CreateFont();
        font.IsBold = Bold;
        font.IsItalic = Italic;
        font.IsStrikeout = Strikeout;
        font.FontName = FontName;
        font.FontHeightInPoints = FontSize;

        cellStyle.SetFont(font);

        if (FillPattern != null)
            cellStyle.FillPattern = FillPattern.Value;
        if (ForegroundColor != null)
            cellStyle.SetFillForegroundColor(ForegroundColor);
        if (BackgroundColor != null)
            cellStyle.SetFillBackgroundColor(BackgroundColor.Value);

        // Alignment settings
        if (HorizontalAlignment != null)
            cellStyle.Alignment = HorizontalAlignment.Value;
        if (VerticalAlignment != null)
            cellStyle.VerticalAlignment = VerticalAlignment.Value;

        cell.CellStyle = cellStyle;
    }

    public static void ApplyManualSizeColumns(ISheet excelSheet, int maxColumnCount)
    {
        for (int columnIndex = 0; columnIndex < maxColumnCount; columnIndex++)
        {
            double maxWidth = 0;

            for (int rowIndex = 0; rowIndex <= excelSheet.LastRowNum; rowIndex++)
            {
                IRow row = excelSheet.GetRow(rowIndex);
                if (row != null)
                {
                    ICell cell = row.GetCell(columnIndex);
                    if (cell != null)
                    {
                        string cellText = cell.StringCellValue;
                        ICellStyle cellStyle = cell.CellStyle;
                        IFont font = excelSheet.Workbook.GetFontAt(cellStyle.FontIndex);
                        int fontSize = font.FontHeightInPoints;

                        // Hücredeki metin uzunluğunu ve yazı tipi boyutunu hesaba kat
                        int cellTextLength = cellText.Length;

                        // Ortalama karakter genişliği için makul bir katsayı kullan
                        double adjustedWidth = cellTextLength * 1.143 * fontSize;

                        // Minimum genişlik ayarı (10 karakter genişliği)
                        if (adjustedWidth < 10 * 256)
                        {
                            adjustedWidth = 10 * 256;
                        }

                        if (adjustedWidth > maxWidth)
                        {
                            maxWidth = adjustedWidth;
                        }
                    }
                }
            }
            // Hesaplanan genişliği ayarla (256 birim = 1 karakter genişliği)
            int finalWidth = (int)Math.Min(maxWidth, 255 * 256); // 255 karakter genişliği sınırını uygula
            excelSheet.SetColumnWidth(columnIndex, finalWidth);
        }
    }
}
