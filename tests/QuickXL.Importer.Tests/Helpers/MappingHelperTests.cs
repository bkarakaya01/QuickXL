using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using QuickXL.Importer.Helpers;
using QuickXL.Importer.Tests.Fixtures;

namespace QuickXL.Importer.Tests.Helpers
{
    [TestFixture]
    public class MappingHelperTests
    {
        [Test]
        public void BuildAttributeMappings_ReturnsCorrectMappings()
        {
            var headers = new[] { "Name", "Age", "Unused" };
            var maps = MappingHelper.BuildAttributeMappings<Person>(headers).ToList();

            // Person fixture’da [XLColumn("Name")] ve [XLColumn("Age")] var
            Assert.That(maps.Count, Is.EqualTo(2));
            Assert.Multiple(() =>
            {
                Assert.That(maps.Any(m => m.Index == 0), Is.True);
                Assert.That(maps.Any(m => m.Index == 1), Is.True);
            });
        }

        [Test]
        public void BuildSetter_ParsesSupportedTypes()
        {
            var piName = typeof(Person).GetProperty(nameof(Person.Name))!;
            var setterName = MappingHelper.BuildSetter(piName);
            var dto1 = new Person();
            setterName(dto1, "Hello");
            Assert.That(dto1.Name, Is.EqualTo("Hello"));

            var piAge = typeof(Person).GetProperty(nameof(Person.Age))!;
            var setterAge = MappingHelper.BuildSetter(piAge);
            var dto2 = new Person();
            setterAge(dto2, "123");
            Assert.That(dto2.Age, Is.EqualTo(123));
        }

        [Test]
        public void BuildSetter_IgnoresInvalidParse()
        {
            var piAge = typeof(Person).GetProperty(nameof(Person.Age))!;
            var setter = MappingHelper.BuildSetter(piAge);
            var p = new Person { Age = 7 };
            setter(p, "notanumber");
            // parse başarısızsa original setter string atandığından p.Age==0
            Assert.That(p.Age, Is.EqualTo(0));
        }
    }
}
