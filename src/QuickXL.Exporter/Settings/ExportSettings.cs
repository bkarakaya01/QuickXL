namespace QuickXL
{
    public class ExportSettings
    {
        public int FirstRowIndex { get; set; }
        public string? SheetName { get; set; }

        public ExportSettings()
        {
            FirstRowIndex = 0;
            SheetName = "QuickXL";
        }
    }
}
