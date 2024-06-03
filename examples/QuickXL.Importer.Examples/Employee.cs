namespace QuickXL.Importer.Examples
{
    public class Employee : IExcelPOCO
    {
        [ExportHeader(nameof(Name))]
        public string Name { get; set; }

        [ExportHeader(nameof(Surname))]
        public string Surname { get; set; }

        [ExportHeader(nameof(Age))]
        public int Age { get; set; }
    }

}
