using QuickXL.Exporter.Tests.SampleProvider;

namespace QuickXL.Exporter.Tests.UnitTests;

[TestFixture]
internal class ExportBuilderTests
{
    [Test]
    public void ExportBuilder_WithData_AddsDataCorrectly()
    {
        // Arrange
        var data = new List<SampleDto> { new SampleDto { Id = 2, Name = "Test2" } };
        var exportBuilder = new ExportBuilder<SampleDto>();

        // Act
        exportBuilder.WithData(data);

        // Assert
        Assert.That(exportBuilder.Data.Count, Is.EqualTo(1));
        Assert.That(exportBuilder.Data[0], Is.EqualTo(data[0]));
    }

    [Test]
    public void ExportBuilder_AddColumn_AddsColumnCorrectly()
    {
        // Arrange
        var exportBuilder = new ExportBuilder<SampleDto>();

        // Act
        exportBuilder.AddColumn("Id", x => x.Id);

        // Assert
        Assert.That(exportBuilder.HeaderPropertySelectors, Has.Count.EqualTo(1));
        Assert.That(exportBuilder.HeaderPropertySelectors.ContainsKey("Id"), Is.True);
    }

    [Test]
    public void Build_ReturnsExporterWithCorrectSettings()
    {
        // Arrange
        var exportBuilder = new ExportBuilder<SampleDto>();

        // Act
        var exporter = exportBuilder.Build(settings =>
        {
            settings.FirstRowIndex = 1;
            settings.SheetName = "TestSheet";
        });

        // Assert
        Assert.That(exporter.WorkbookSettings.FirstRowIndex, Is.EqualTo(1));
        Assert.That(exporter.WorkbookSettings.SheetName, Is.EqualTo("TestSheet"));
    }
}
