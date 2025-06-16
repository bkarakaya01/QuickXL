using NUnit.Framework;
using QuickXL.Exporter.Tests.SampleProvider;

namespace QuickXL.Exporter.Tests.UnitTests;

[TestFixture]
public class XLSheetTests
{
    //private XLSheet<SampleDto> xlSheet;

    //[SetUp]
    //public void Setup()
    //{
    //    xlSheet = new XLSheet<SampleDto>();
    //}

    //[Test]
    //public void AddCell_AddsCellCorrectly()
    //{
    //    // Arrange
    //    int rowIndex = 1;
    //    int columnIndex = 1;
    //    string value = "TestValue";

    //    // Act
    //    xlSheet.AddCell(rowIndex, columnIndex, value);

    //    // Assert
    //    var cell = xlSheet[rowIndex, columnIndex];
    //    Assert.That(cell, Is.Not.Null);
    //    Assert.That(cell.Value, Is.EqualTo(value));
    //}

    //[Test]
    //public void AddCell_WhenCellAlreadyExists_ThrowsException()
    //{
    //    // Arrange
    //    int rowIndex = 1;
    //    int columnIndex = 1;
    //    string value = "TestValue";

    //    // Act
    //    xlSheet.AddCell(rowIndex, columnIndex, value);

    //    // Assert
    //    Assert.Throws<InvalidOperationException>(() => xlSheet.AddCell(rowIndex, columnIndex, value));
    //}

    //[Test]
    //public void AddHeader_AddsHeaderCorrectly()
    //{
    //    // Arrange
    //    int columnIndex = 0;
    //    string header = "Header";

    //    // Act
    //    xlSheet.AddHeader(columnIndex, header);

    //    // Assert
    //    var cell = xlSheet[xlSheet.FirstRowIndex, columnIndex];
    //    Assert.That(cell, Is.Not.Null);
    //    Assert.That(cell.Value, Is.EqualTo(header));
    //}

    //[Test]
    //public void GetLastRow_ReturnsCorrectRowIndex()
    //{
    //    // Arrange
    //    xlSheet.AddCell(1, 0, "Test1");
    //    xlSheet.AddCell(2, 0, "Test2");

    //    // Act
    //    int lastRowIndex = xlSheet.GetLastRow();

    //    // Assert
    //    Assert.That(lastRowIndex, Is.EqualTo(2));
    //}

    //[Test]
    //public void GetLastColumn_ReturnsCorrectColumnIndex()
    //{
    //    // Arrange
    //    xlSheet.AddCell(0, 1, "Test1");
    //    xlSheet.AddCell(0, 2, "Test2");

    //    // Act
    //    int lastColumnIndex = xlSheet.GetLastColumn(0);

    //    // Assert
    //    Assert.That(lastColumnIndex, Is.EqualTo(2));
    //}
}
