namespace QuickXL.Core.Contracts.Settings
{
    public interface IWorkbookSettings
    {
        int FirstRowIndex { get; set; }
        string? SheetName { get; set; }
    }
}
