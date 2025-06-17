namespace QuickXL.Core.Settings;

internal record WorkbookSettings
{
    public int FirstRowIndex { get; set; } = 0;
    public string SheetName { get; set; } = null!;
}
