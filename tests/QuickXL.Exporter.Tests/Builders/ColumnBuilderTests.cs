using QuickXL.Exporter.Tests.Fixtures;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace QuickXL.Exporter.Tests.Builders
{
    [TestFixture]
    public class ColumnBuilderTests
    {
        private ExportBuilder<Person> _exportBuilder = null!;
        private ColumnBuilder<Person> _columnBuilder = null!;

        [SetUp]
        public void SetUp()
        {
            var settings = new WorkbookSettings { SheetName = "S", FirstRowIndex = 0 };
            _exportBuilder = new ExportBuilder<Person>(settings);
            _columnBuilder = _exportBuilder.ColumnBuilder;
        }

        [Test]
        public void AddColumn_WithValidSelector_AddsItem()
        {
            _columnBuilder.AddColumn(x => x.Name);
            Assert.That(_columnBuilder.ColumnBuilderItems.Count, Is.EqualTo(1));

            var item = _columnBuilder.ColumnBuilderItems.Single();
            Assert.That(item.HeaderName, Is.EqualTo("Name"));
        }

        [Test]
        public void AddColumn_WithCustomHeaderName_OverridesDefault()
        {
            _columnBuilder.AddColumn(x => x.Age, opts => opts.HeaderName = "Years");
            var item = _columnBuilder.ColumnBuilderItems.Single();
            Assert.That(item.HeaderName, Is.EqualTo("Years"));
        }
    }
}
