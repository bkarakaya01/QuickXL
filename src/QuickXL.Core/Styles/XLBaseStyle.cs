using Ardalis.GuardClauses;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using QuickXL.Core.Models.Cells;
using QuickXL.Core.Models.Colors;

namespace QuickXL.Core.Styles;

public abstract class XLBaseStyle
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

    public XLBaseStyle()
    {
        Bold = false;
        Italic = false;
        Strikeout = false;
        FontName = "Calibri";
        FontSize = 12;
    }
    internal virtual void Apply(ISheet excelSheet, XLCell xlCell)
    {
        ICell? cell = xlCell.Cell;

        Guard.Against.Null(excelSheet);
        Guard.Against.Null(cell);

        if (excelSheet.Workbook.CreateCellStyle() is not XSSFCellStyle cellStyle)
            throw new NullReferenceException(nameof(cellStyle));

        // Font settings
        var font = excelSheet.Workbook.CreateFont();
        font.IsBold = Bold;
        font.IsItalic = Italic;
        font.IsStrikeout = Strikeout;
        font.FontName = FontName;
        font.FontHeightInPoints = FontSize;

        cellStyle.SetFont(font);

        // Color settings
        if (ForegroundColor != null)
            cellStyle.SetFillForegroundColor(ForegroundColor);
        if (BackgroundColor != null)
            cellStyle.SetFillBackgroundColor(BackgroundColor);
        if (FillPattern != null)
            cellStyle.FillPattern = FillPattern.Value;

        // Alignment settings
        if (HorizontalAlignment != null)
            cellStyle.Alignment = HorizontalAlignment.Value;
        if (VerticalAlignment != null)
            cellStyle.VerticalAlignment = VerticalAlignment.Value;

        cell.CellStyle = cellStyle;
    }
}
