namespace QuickXL.Exporter.Tests.UnitTests;

[TestFixture]
[TestFixture]
public class ExporterTests
{
    //private IExporter<SampleDto> exporter;
    //private ExportBuilder<SampleDto> exportBuilder;

    //[SetUp]
    //public void Setup()
    //{
    //    exportBuilder = new ExportBuilder<SampleDto>()
    //        .WithData(SampleDataProvider.Create())
    //        .AddColumn("Id", x => x.Id)
    //        .AddColumn("Name", x => x.Name);

    //    exporter = exportBuilder.Build(settings =>
    //    {
    //        settings.FirstRowIndex = 1;
    //        settings.SheetName = "TestSheet";
    //    });
    //}

    //[Test]
    //public void Export_WithValidData_ReturnsSuccessResult()
    //{
    //    // Act
    //    var result = exporter.Export();

    //    // Assert
    //    Assert.That(result, Is.Not.Null);
    //    Assert.Multiple(() =>
    //    {
    //        Assert.That(result.Succeeded, Is.True);
    //        Assert.That(result.Data, Is.Not.Null);
    //    });
    //}

    //[Test]
    //public void Export_WhenExportBuilderIsNull_ThrowsException()
    //{
    //    ExportBuilder<SampleDto> temp = exporter.ExportBuilder;

    //    // Arrange
    //    exporter.ExportBuilder = null;

    //    // Act
    //    var result = exporter.Export();

    //    // Assert
    //    Assert.That(result, Is.Not.Null);
    //    Assert.Multiple(() =>
    //    {
    //        Assert.That(result.Succeeded, Is.False);
    //        Assert.That(result.Errors, Is.Not.Null);
    //    });
    //    Assert.That(result.Errors, Has.One.Items);
    //    Assert.That(result.Errors[0].Code, Is.EqualTo(typeof(ArgumentNullException).Name));

    //    exporter.ExportBuilder = temp;
    //}

    //[Test]
    //public void Export_WhenWorkbookSettingsIsNull_ThrowsException()
    //{
    //    ExportSettings temp = exporter.WorkbookSettings;

    //    // Arrange
    //    exporter.WorkbookSettings = null;

    //    // Act
    //    var result = exporter.Export();

    //    // Assert
    //    Assert.That(result, Is.Not.Null);
    //    Assert.Multiple(() =>
    //    {
    //        Assert.That(result.Succeeded, Is.False);
    //        Assert.That(result.Errors, Is.Not.Null);
    //    });
    //    Assert.That(result.Errors, Has.One.Items);
    //    Assert.That(result.Errors[0].Code, Is.EqualTo(typeof(ArgumentNullException).Name));

    //    exporter.WorkbookSettings = temp;
    //}
}