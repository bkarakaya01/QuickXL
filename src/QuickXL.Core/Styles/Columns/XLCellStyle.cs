using QuickXL.Core.Models.Colors;

namespace QuickXL.Core.Styles.Columns;

public sealed class XLCellStyle
{
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public bool Strikeout { get; set; }
    public string FontName { get; set; } = "Calibri";
    public short FontSize { get; set; } = 12;
    public XLColor? ForegroundColor { get; set; }
    public XLColor? BackgroundColor { get; set; }
}
