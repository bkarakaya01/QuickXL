namespace QuickXL
{
    public interface IExcelSheet
    {
        void AddCell(int rowIndex, int columnIndex, string value);
        void AddHeader(int columnIndex, string header);
    }
}
