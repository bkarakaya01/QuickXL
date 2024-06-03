using QuickXL;
using QuickXL.Importer.Examples;


WorkbookSettings workbookSettings = new();

var data = PopulateData();

var result = Exporter.Export(workbookSettings, data);

var path = @"C:\Users\TeknikMedya\Projects\QuickXL_Export.xlsx";

if(result.Success)
{
    using var fs = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite);

    result.Data!.CopyTo(fs);
}

static List<Employee> PopulateData()
{
    var employees = new List<Employee>();

    for (int index = 0; index < 1000; index++)
    {
        employees.Add(new()
        {
            Name = $"Name_{index}",
            Surname = $"Surname_{index}",
            Age = index + 1,
        });
    }

    return employees;
}

