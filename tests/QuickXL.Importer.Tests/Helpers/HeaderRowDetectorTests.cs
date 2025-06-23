using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Importer.Helpers;

namespace QuickXL.Importer.Tests.Helpers
{
    [TestFixture]
    public class HeaderRowDetectorTests
    {
        private static Row MakeRow(int rowIndex, params string[] values)
        {
            var row = new Row { RowIndex = (uint)rowIndex };
            foreach (var text in values)
            {
                row.Append(new Cell
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue(text)
                });
            }
            return row;
        }

        [Test]
        public void Detect_UsesManualSetting_WhenProvided()
        {
            var rows = new List<Row>
            {
                MakeRow(0, "X"),
                MakeRow(1, "Y")
            };
            var settings = new HeaderRowSettings { StartsAt = 5 };
            int idx = HeaderRowDetector.Detect<object>(
                rows,
                settings,
                new[] { "DoesNotMatter" }
            );
            Assert.That(idx, Is.EqualTo(5));
        }

        [Test]
        public void Detect_PicksRowWithMostMatches()
        {
            // Row0: [A,B]  → 2 matches
            // Row1: [X,A]  → 1 match
            // Row2: [B,C]  → 1 match
            var expectedHeaders = new[] { "A", "B" };
            var rows = new List<Row>
            {
                MakeRow(0, "A", "B"),
                MakeRow(1, "X", "A"),
                MakeRow(2, "B", "C")
            };

            int idx = HeaderRowDetector.Detect<object>(
                rows,
                settings: null,
                expectedHeaders
            );

            // Row0 has the highest score (2), so should be chosen
            Assert.That(idx, Is.EqualTo(0));
        }
    }
}
