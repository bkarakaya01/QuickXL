using NUnit.Framework;
using QuickXL.Exporter.Helpers;
using QuickXL.Exporter.Tests.Fixtures;
using System.Collections.Generic;

namespace QuickXL.Exporter.Tests.Helpers
{
    [TestFixture]
    public class ColumnWidthCalculatorTests
    {
        [Test]
        public void Calculate_ConsidersHeaderAndDataLengths()
        {
            var list = new List<Person>
            {
                new Person { Name = "A", Age = 123 },
                new Person { Name = "LongName", Age = 4 }
            };

            // header length = 4 ("Name")
            double width = ColumnWidthCalculator.Calculate(list, p => p.Name, headerLength: 4);
            Assert.That(width, Is.GreaterThanOrEqualTo(4 * 1.2));
            Assert.That(width, Is.LessThanOrEqualTo(255));
        }
    }
}
