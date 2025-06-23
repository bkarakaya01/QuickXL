using DocumentFormat.OpenXml.Packaging;
using QuickXL.Exporter.Helpers;
using QuickXL.Exporter.Tests.Fixtures;

namespace QuickXL.Exporter.Tests.Helpers
{
    [TestFixture]
    public class OpenXmlWorkbookHelperTests
    {
        [Test]
        public void CreateWorkbook_GeneratesValidXlsx_WithHeadersAndValues()
        {
            var settings = new WorkbookSettings { SheetName = "Sheet1", FirstRowIndex = 0 };
            var builder = new ExportBuilder<Person>(settings)
                .WithData(
                [
                    new() { Name = "X", Age = 1, IsAdmin = true }
                ])
            .AddColumn(x => x.Name)
            .AddColumn(x => x.Age)
            .AddColumn(x => x.IsAdmin);

            var helper = new OpenXmlWorkbookHelper<Person>();
            byte[] bytes = helper.CreateWorkbook(builder.ExportBuilder);

            // Open the workbook in-memory and verify
            using var ms = new MemoryStream(bytes);
            using var doc = SpreadsheetDocument.Open(ms, false);
            var sheet = doc.WorkbookPart!.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().First();
            var wsPart = (WorksheetPart)doc.WorkbookPart!.GetPartById(sheet.Id!);
            var rows = wsPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()!.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().ToList();

            // Header + 1 data row
            Assert.That(rows.Count, Is.EqualTo(2));

            var headerCells = rows[0].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ToList();
            Assert.That(headerCells.Select(c => c.CellValue!.Text), Is.EqualTo(new[] { "Name", "Age", "IsAdmin" }));

            var dataCells = rows[1].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().ToList();
            
            Assert.Multiple(() =>
            {
                Assert.That(dataCells[0].CellValue!.Text, Is.EqualTo("X"));
                Assert.That(dataCells[1].CellValue!.Text, Is.EqualTo("1"));
                Assert.That(dataCells[2].CellValue!.Text, Is.EqualTo("True"));
            });
        }
    }
}
