namespace QuickXL.Exporter;

public sealed class XLCellStyle
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Strikeout { get; set; }
    public string FontName { get; set; } = "Calibri";
    public short FontSize { get; set; } = 12;
    public XLColor? ForegroundColor { get; set; }
    public XLColor? BackgroundColor { get; set; }
    public XLHorizontalAlignment? HorizontalAlignment { get; set; }
    public XLVerticalAlignment? VerticalAlignment { get; set; }
}
