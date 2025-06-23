using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using QuickXL.Importer.Helpers;
using QuickXL.Importer.Tests.Builders;
using QuickXL.Importer.Tests.Fixtures;
using System.IO;
using System.Linq;

namespace QuickXL.Importer.Tests.Helpers
{
    [TestFixture]
    public class WorksheetReaderTests
    {
        private byte[] _workbookBytes = null!;

        [SetUp]
        public void SetUp()
        {
            // Basit bir xlsx oluştur: 1 sheet, 2 satır, 2 sütun
            _workbookBytes = ImportBuilderTests.CreateTestWorkbook(
                new[] { "Col1", "Col2" },
                new[] { new[] { "A1", "B1" }, new[] { "A2", "B2" } }
            );
        }

        [Test]
        public void ReadRows_ReturnsAllRows()
        {
            using var ms = new MemoryStream(_workbookBytes);
            var rows = WorksheetReader.ReadRows(ms, out var wbPart);

            Assert.That(rows.Count, Is.EqualTo(3), "Header + 2 data rows");
            // İlk hücrenin değeri doğru okunuyor mu?
            var firstCell = rows[0].Elements<Cell>().First();
            Assert.That(WorksheetReader.GetCellValue(firstCell, wbPart), Is.EqualTo("Col1"));
        }

        [Test]
        public void GetCellValue_SharedStringAndInlineText()
        {
            // SharedString test: header row already uses SharedString in our helper workbook
            using var ms = new MemoryStream(_workbookBytes);
            var rows = WorksheetReader.ReadRows(ms, out var wbPart);

            // İlk hücre shared string olarak okunmalı
            var headerCell = rows[0].Elements<Cell>().First();
            var text = WorksheetReader.GetCellValue(headerCell, wbPart);
            Assert.That(text, Is.EqualTo("Col1"));

            // Bir veri hücresi inline
            var dataCell = rows[1].Elements<Cell>().First();
            Assert.That(WorksheetReader.GetCellValue(dataCell, wbPart), Is.EqualTo("A1"));
        }
    }
}
