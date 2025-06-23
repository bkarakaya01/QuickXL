using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using QuickXL.Importer.Exceptions;
using QuickXL.Importer.Tests.Fixtures;

namespace QuickXL.Importer.Tests.Builders
{
    [TestFixture]
    public class ImportBuilderTests
    {
        private ImportSettings _settings = null!;
        private ImportBuilder<Person> _builder = null!;
        private byte[] _workbookBytes = null!;

        [SetUp]
        public void SetUp()
        {
            _settings = new ImportSettings();
            _builder = new ImportBuilder<Person>(_settings);

            // Create an in-memory Excel workbook with header + one data row
            var headers = new[] { "Name", "Age" };
            var dataRow = new[] { "Alice", "42" };
            _workbookBytes = CreateTestWorkbook(headers, new[] { dataRow });
        }

        /// <summary>
        /// Helper to create a minimal in-memory .xlsx for tests, with one sheet,
        /// a header row, and arbitrary data rows.
        /// </summary>
        internal static byte[] CreateTestWorkbook(string[] headers, string[][] rows)
        {
            using var ms = new MemoryStream();
            // Create the spreadsheet package
            using (var document = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, false))
            {
                // 1) Workbook + Styles
                var wbPart = document.AddWorkbookPart();
                wbPart.Workbook = new Workbook();
                var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
                stylesPart.Stylesheet = new Stylesheet();
                stylesPart.Stylesheet.Save();

                // 2) Sheets container + single sheet
                var sheets = wbPart.Workbook.AppendChild(new Sheets());
                var wsPart = wbPart.AddNewPart<WorksheetPart>();
                sheets.Append(new Sheet
                {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1,
                    Name = "Sheet1"
                });

                // 3) Worksheet + SheetData
                wsPart.Worksheet = new Worksheet();
                var sheetData = wsPart.Worksheet.AppendChild(new SheetData());

                // 4) Header row at index 1
                var headerRow = new Row { RowIndex = 1U };
                for (int c = 0; c < headers.Length; c++)
                {
                    headerRow.Append(new Cell
                    {
                        CellReference = GetCellRef(c, 1),
                        DataType = CellValues.String,
                        CellValue = new CellValue(headers[c])
                    });
                }
                sheetData.Append(headerRow);

                // 5) Data rows starting at index 2
                for (int r = 0; r < rows.Length; r++)
                {
                    var dataRow = new Row { RowIndex = (uint)(r + 2) };
                    for (int c = 0; c < rows[r].Length; c++)
                    {
                        dataRow.Append(new Cell
                        {
                            CellReference = GetCellRef(c, r + 2),
                            DataType = CellValues.String,
                            CellValue = new CellValue(rows[r][c])
                        });
                    }
                    sheetData.Append(dataRow);
                }

                // 6) Save all parts
                wsPart.Worksheet.Save();
                wbPart.Workbook.Save();  // Ensure the package writes to the stream
            }

            // 7) Return the underlying byte array
            return ms.ToArray();

            // Helper to convert zero-based column & row to "A1" style reference
            static string GetCellRef(int colIndex, int rowIndex)
            {
                const string Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                string colName = "";
                int dividend = colIndex + 1;
                while (dividend > 0)
                {
                    int mod = (dividend - 1) % 26;
                    colName = Letters[mod] + colName;
                    dividend = (dividend - mod) / 26;
                }
                return $"{colName}{rowIndex}";
            }
        }

        [Test]
        public void FromStream_SetsStreamProperly()
        {
            using var ms = new MemoryStream(_workbookBytes);
            _builder.FromStream(ms);
            Assert.That(_builder.Stream, Is.SameAs(ms));
        }

        [Test]
        public void Import_AppliesAttributeMappings_ByDefault()
        {
            // Arrange: decorate Person properties with [XLColumn]
            using var ms = new MemoryStream(_workbookBytes);
            _builder.FromStream(ms);

            // Act
            var result = _builder.Import();

            // Assert
            Assert.That(result, Has.Count.EqualTo(1));
            var person = result[0];
            Assert.Multiple(() =>
            {
                Assert.That(person.Name, Is.EqualTo("Alice"));
                Assert.That(person.Age, Is.EqualTo(42));
            });
        }

        [Test]
        public void Import_RespectsManualMapping_WhenUseAttributesDisabled()
        {
            // Arrange: disable attribute mapping and map manually
            _settings.UseAttributes = false;
            using var ms = new MemoryStream(_workbookBytes);
            _builder.FromStream(ms)
                    .Map(x => x.Name, 0)
                    .Map(x => x.Age, 1);

            // Act
            var result = _builder.Import();

            // Assert
            Assert.That(result, Has.Count.EqualTo(1));
            var person = result[0];
            
            Assert.Multiple(() =>
            {
                Assert.That(person.Name, Is.EqualTo("Alice"));
                Assert.That(person.Age, Is.EqualTo(42));
            });
        }

        [Test]
        public void Import_ThrowsIfNoSourceSpecified()
        {
            var ex = Assert.Throws<NoSourceSpecifiedException>(() => _builder.Import());
        }
    }
}