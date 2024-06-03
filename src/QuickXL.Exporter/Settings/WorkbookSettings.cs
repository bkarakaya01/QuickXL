namespace QuickXL
{
    public class WorkbookSettings
    {
        public int FirstRowIndex { get; set; }
        public string? SheetName { get; set; }

        public WorkbookSettings()
        {
            FirstRowIndex = 0;
            SheetName = "QuickXL";
        }
    }
}
