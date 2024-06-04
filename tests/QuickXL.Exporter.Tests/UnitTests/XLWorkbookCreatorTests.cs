using NUnit.Framework;
using QuickXL.Exporter.Tests.SampleProvider;

namespace QuickXL.Exporter.Tests.UnitTests;

[TestFixture]
internal class XLWorkbookCreatorTests
{
    private Exporter<SampleDto> exporter;
    private ExportBuilder<SampleDto> exportBuilder;

    [SetUp]
    public void Setup()
    {
        exportBuilder = new ExportBuilder<SampleDto>()
            .WithData(SampleDataProvider.Create())
            .AddColumn("Id", x => x.Id)
            .AddColumn("Name", x => x.Name);

        exporter = exportBuilder.Build(settings =>
        {
            settings.FirstRowIndex = 1;
            settings.SheetName = "TestSheet";
        });
    }

    [Test]
    public void CreateWorkbook_CreatesWorkbookWithCorrectSheet()
    {
        // Arrange
        var workbookCreator = new XLWorkbookCreator<SampleDto>(exporter);

        // Act
        var workbook = workbookCreator.CreateWorkbook();

        // Assert
        Assert.That(workbook, Is.Not.Null);
        var sheet = workbook.GetSheet("TestSheet");
        Assert.That(sheet, Is.Not.Null);
    }

    [Test]
    public void CreateWorkbook_CreatesWorkbookWithCorrectHeaders()
    {
        // Arrange
        var workbookCreator = new XLWorkbookCreator<SampleDto>(exporter);

        // Act
        var workbook = workbookCreator.CreateWorkbook();
        var sheet = workbook.GetSheet("TestSheet");
        var headerRow = sheet.GetRow(1);

        // Assert
        Assert.That(headerRow, Is.Not.Null);
        Assert.That(headerRow.GetCell(0).StringCellValue, Is.EqualTo("Id"));
        Assert.That(headerRow.GetCell(1).StringCellValue, Is.EqualTo("Name"));
    }

    [Test]
    public void CreateWorkbook_CreatesWorkbookWithCorrectData()
    {
        // Arrange
        var workbookCreator = new XLWorkbookCreator<SampleDto>(exporter);

        // Act
        var workbook = workbookCreator.CreateWorkbook();
        var sheet = workbook.GetSheet("TestSheet");

        for (int i = 0; i < SampleDataProvider.Create().Count; i++)
        {
            var dataRow = sheet.GetRow(i + 2); // this.exporter.WorkbookSettings.FirstRowIndex is set to 1, so data starts from index 2

            // Assert
            Assert.That(dataRow, Is.Not.Null);
            Assert.That(dataRow.GetCell(0).StringCellValue, Is.EqualTo(SampleDataProvider.Create()[i].Id.ToString()));
            Assert.That(dataRow.GetCell(1).StringCellValue, Is.EqualTo(SampleDataProvider.Create()[i].Name));
        }
    }
}
