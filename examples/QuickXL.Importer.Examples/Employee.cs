namespace QuickXL.Importer.Examples
{
    public class Employee : IExcelPOCO
    {
        [XLHeader(nameof(Name))]
        public string Name { get; set; }

        [XLHeader(nameof(Surname))]
        public string Surname { get; set; }

        [XLHeader(nameof(Age))]
        public int Age { get; set; }
    }

}
