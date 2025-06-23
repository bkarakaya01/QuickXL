namespace QuickXL.Importer.Tests.Fixtures
{
    public class Person
    {
        [XLColumn("Name")]
        public string Name { get; set; } = string.Empty;

        [XLColumn("Age")]
        public int Age { get; set; }
    }
}
