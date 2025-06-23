using QuickXL.Exporter.Helpers;

namespace QuickXL.Exporter.Tests.Helpers
{
    [TestFixture]
    public class ExcelReferenceHelperTests
    {
        [TestCase(0, "A")]
        [TestCase(1, "B")]
        [TestCase(25, "Z")]
        [TestCase(26, "AA")]
        [TestCase(27, "AB")]
        [TestCase(701, "ZZ")]
        [TestCase(702, "AAA")]
        public void GetName_ReturnsExpected(int index, string expected)
        {
            Assert.That(ExcelReferenceHelper.GetColumnName(index), Is.EqualTo(expected));
        }
    }
}
