using QuickXL.Exporter.Styles.Columns;

namespace QuickXL.Exporter.Styles
{
    public class XLGeneralStyle
    {
        public bool AutoSizeColumns { get; set; }

        public XLCellStyle HeaderStyle { get; set; }    
        public XLCellStyle CellStyle { get; set; }

        public XLGeneralStyle()
        {
            HeaderStyle = new();
            CellStyle = new();
        }
    }
}
