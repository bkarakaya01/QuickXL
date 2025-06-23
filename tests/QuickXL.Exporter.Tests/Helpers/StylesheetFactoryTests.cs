using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Exporter.Helpers;
using QuickXL.Exporter.Styles;  // for StylesheetFactory
using QuickXL.Exporter.Tests.Fixtures;

namespace QuickXL.Exporter.Tests.Helpers
{
    [TestFixture]
    public class StylesheetFactoryTests
    {
        private WorkbookSettings _settings = null!;
        private ExportBuilder<Person> _builder = null!;

        [SetUp]
        public void Setup()
        {
            _settings = new WorkbookSettings
            {
                SheetName = "Sheet1",
                FirstRowIndex = 0
            };
            _builder = new ExportBuilder<Person>(_settings);
        }

        [Test]
        public void Create_DefaultBuilder_ProducesCorrectStructure()
        {
            // arrange
            _builder.ColumnBuilder.AddColumn(x => x.Name);
            _builder.ColumnBuilder.AddColumn(x => x.Age);

            // act
            Stylesheet sheet = StylesheetFactory.Create(_builder);

            // assert counts
            Assert.That(sheet.Fonts.Count(), Is.EqualTo(2), "should have default + header fonts");
            Assert.That(sheet.Fills.Count(), Is.EqualTo(2), "should have two fills");
            Assert.That(sheet.Borders.Count(), Is.EqualTo(1), "should have one border");
            Assert.That(sheet.CellStyleFormats.Count(), Is.EqualTo(1), "should have one style format");
            Assert.That(sheet.CellFormats.Count(), Is.EqualTo(3), "should have default, header, data formats");
        }

        [Test]
        public void Create_WithGeneralStyle_AppliesCustomHeaderFont()
        {
            // arrange
            var gs = new XLGeneralStyle
            {
                HeaderStyle = new XLCellStyle
                {
                    FontSize = 14,
                    FontName = "Arial",
                    ForegroundColor = XLColor.Make(255, 0, 0) // red
                }
            };
            _builder.ColumnBuilder.AddGeneralStyle(gs);
            _builder.ColumnBuilder.AddColumn(x => x.Name);

            // act
            Stylesheet sheet = StylesheetFactory.Create(_builder);

            // assert header font matches
            var headerFont = sheet.Fonts.Elements<Font>().ElementAt(1);
            var sizeNode = headerFont.Elements<FontSize>().Single();
            var nameNode = headerFont.Elements<FontName>().Single();
            var colorNode = headerFont.Elements<Color>().Single();

            Assert.That(sizeNode.Val, Is.EqualTo(14d), "Header font size should match");
            Assert.That(nameNode.Val.ToString(), Is.EqualTo("Arial"), "Header font name should match");
            Assert.That(colorNode.Rgb.Value, Is.EqualTo("FF0000"), "Header font color should match red");
        }
    }
}
