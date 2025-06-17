namespace QuickXL.Core.Settings;

internal record WorkbookSettings
{
    public int FirstRowIndex { get; set; }
    public string? SheetName { get; set; }
}
