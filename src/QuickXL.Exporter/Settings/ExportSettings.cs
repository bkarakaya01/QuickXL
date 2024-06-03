
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace QuickXL
{
    /// <summary>
    /// Settings class to manage the created Excel Workbook's style.
    /// </summary>
    public class ExportSettings
    {
        /// <summary>
        /// Decides whether columns should auto sized or not.
        /// </summary>
        public bool AutoSizeColumns { get; set; } = false;

        /// <summary>
        /// Decides whether header row's cells should be bold or not.
        /// </summary>
        public bool BoldHeader { get; set; } = false;

        /// <summary>
        /// Decides whether header row's cells should be italic or not.
        /// </summary>
        public bool ItalicHeader { get; set; } = false;

        /// <summary>
        /// Decides whether header row's cells should be strikeout or not.
        /// </summary>
        public bool StrikeoutHeader { get; set; } = false;

        /// <summary>
        /// Font name pointing the font which will be used for header cells.
        /// </summary>
        public string HeaderFontName { get; set; } = "Calibri";

        /// <summary>
        /// Font size of the header cells.
        /// </summary>
        public short HeaderFontSize { get; set; } = 12;

        /// <summary>
        /// Color settings which will be used for header cells fill foreground color.
        /// </summary>
        public XSSFColor? HeaderForegroundColor { get; set; }

        /// <summary>
        /// Color settings which will be used for header cells fill background color.
        /// </summary>
        public XSSFColor? HeaderBackgroundColor { get; set; }

        /// <summary>
        /// Fill pattern settings which will be used for header cells style.
        /// </summary>
        public FillPattern? HeaderFillPattern { get; set; }

        /// <summary>
        /// Horizontal Alignment settings for the header cells.
        /// </summary>
        public HorizontalAlignment? HeaderHorizontalAlignment { get; set; }

        /// <summary>
        /// Vertical Alignment settings for the header cells.
        /// </summary>
        public VerticalAlignment? HeaderVerticalAlignment { get; set; }

        /// <summary>
        /// Decides whether value row's columns should be bold or not.
        /// </summary>
        public bool BoldValueCell { get; set; } = false;

        /// <summary>
        /// Decides whether value row's columns should be italic or not.
        /// </summary>
        public bool ItalicValueCell { get; set; } = false;

        /// <summary>
        /// Decides whether value row's columns should be strikeout or not.
        /// </summary>
        public bool StrikeoutValueCell { get; set; } = false;

        /// <summary>
        /// Font name pointing the font which will be used for value columns.
        /// </summary>
        public string ValueCellFontName { get; set; } = "Calibri";

        /// <summary>
        /// Font size of the value cells.
        /// </summary>
        public short ValueCellFontSize { get; set; } = 11;

        /// <summary>
        /// Color settings which will be used for value columns fill foreground color.
        /// </summary>
        public XSSFColor? ValueCellForegroundColor { get; set; }

        /// <summary>
        /// Color settings which will be used for value columns fill background color.
        /// </summary>
        public XSSFColor? ValueCellBackgroundColor { get; set; }

        /// <summary>
        /// Fill pattern settings which will be used for value cell style.
        /// </summary>
        public FillPattern? ValueCellFillPattern { get; set; }

        /// <summary>
        /// Horizontal Alignment settings for the value rows.
        /// </summary>
        public HorizontalAlignment? ValueCellHorizontalAlignment { get; set; }

        /// <summary>
        /// Vertical Alignment settings for the value rows.
        /// </summary>
        public VerticalAlignment? ValueCellVerticalAlignment { get; set; }

        public void ApplyHeaderSettings(ISheet excelSheet, ICell headerCell)
        {
            if (excelSheet == null)
                throw new ArgumentNullException(nameof(excelSheet));

            if (headerCell == null)
                throw new ArgumentNullException(nameof(headerCell));

            if (excelSheet.Workbook.CreateCellStyle() is not XSSFCellStyle cellStyle)
                throw new NullReferenceException(nameof(cellStyle));

            //Font settings
            var font = excelSheet.Workbook.CreateFont();
            font.IsBold = BoldHeader;
            font.IsItalic = ItalicHeader;
            font.IsStrikeout = StrikeoutHeader;
            font.FontName = HeaderFontName;
            font.FontHeightInPoints = HeaderFontSize;

            cellStyle.SetFont(font);

            //Color settings
            if (HeaderForegroundColor != null)
                cellStyle.SetFillForegroundColor(HeaderForegroundColor);
            if (HeaderBackgroundColor != null)
                cellStyle.SetFillBackgroundColor(HeaderBackgroundColor);
            if (HeaderFillPattern != null)
                cellStyle.FillPattern = HeaderFillPattern.Value;

            //Alignment settings
            if (HeaderHorizontalAlignment != null)
                cellStyle.Alignment = HeaderHorizontalAlignment.Value;
            if (HeaderVerticalAlignment != null)
                cellStyle.VerticalAlignment = HeaderVerticalAlignment.Value;

            headerCell.CellStyle = cellStyle;
        }

        public void ApplyValueCellSettings(ISheet excelSheet, ICell valueCell)
        {
            if (excelSheet == null)
                throw new ArgumentNullException(nameof(excelSheet));

            if (valueCell == null)
                throw new ArgumentNullException(nameof(valueCell));

            if (excelSheet.Workbook.CreateCellStyle() is not XSSFCellStyle cellStyle)
                throw new NullReferenceException(nameof(cellStyle));

            //Font settings
            var font = excelSheet.Workbook.CreateFont();
            font.IsBold = BoldValueCell;
            font.IsItalic = ItalicValueCell;
            font.IsStrikeout = StrikeoutValueCell;
            font.FontName = ValueCellFontName;
            font.FontHeightInPoints = ValueCellFontSize;

            cellStyle.SetFont(font);

            //Color settings
            if (ValueCellForegroundColor != null)
                cellStyle.SetFillForegroundColor(ValueCellForegroundColor);
            if (ValueCellBackgroundColor != null)
                cellStyle.SetFillBackgroundColor(ValueCellBackgroundColor);
            if (ValueCellFillPattern != null)
                cellStyle.FillPattern = ValueCellFillPattern.Value;

            //Alignment settings
            if (ValueCellHorizontalAlignment != null)
                cellStyle.Alignment = ValueCellHorizontalAlignment.Value;
            if (ValueCellVerticalAlignment != null)
                cellStyle.VerticalAlignment = ValueCellVerticalAlignment.Value;

            valueCell.CellStyle = cellStyle;
        }

        public void ApplyAutoSizeColumns(ISheet excelSheet, int maxColumnCount)
        {
            if (AutoSizeColumns)
                for (int columnIndex = 0; columnIndex < maxColumnCount; columnIndex++)
                    excelSheet.AutoSizeColumn(columnIndex);
        }
    }
}
