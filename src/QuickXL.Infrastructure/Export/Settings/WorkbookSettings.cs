using QuickXL.Core.Contracts.Settings;

namespace QuickXL.Infrastructure.Export.Settings
{
    public class WorkbookSettings : IWorkbookSettings
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
