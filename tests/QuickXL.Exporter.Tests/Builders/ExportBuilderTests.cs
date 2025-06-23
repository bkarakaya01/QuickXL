using QuickXL.Exporter.Tests.Fixtures;

namespace QuickXL.Exporter.Tests.Builders
{
    [TestFixture]
    public class ExportBuilderTests
    {
        private ExportBuilder<Person> _builder = null!;

        [SetUp]
        public void SetUp()
        {
            var settings = new WorkbookSettings { SheetName = "S", FirstRowIndex = 1 };
            _builder = new ExportBuilder<Person>(settings);
        }

        [Test]
        public void WithData_AddsAllItems()
        {
            var list = new List<Person>
            {
                new Person { Name = "A", Age = 10 },
                new Person { Name = "B", Age = 20 }
            };

            _builder.WithData(list);

            Assert.That(_builder.Data, Has.Count.EqualTo(2));
            Assert.That(_builder.ColumnBuilder, Is.Not.Null);
        }
    }
}
