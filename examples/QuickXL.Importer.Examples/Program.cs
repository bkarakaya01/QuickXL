using QuickXL;
using QuickXL.Importer.Examples;


ExportStyleSheet workbookSettings = new();

var data = PopulateData();

var builder = new ExportBuilder<Employee>()
    .WithData(data)
    .AddColumn("Name", x => x.Name)
    .AddColumn("Surname", x => x.Surname)
    .Build(cfg =>
    {
        cfg.SheetName = "Test_1234";
        cfg.FirstRowIndex = 0;
    });

var result = builder.Export();

string path = @"C:\Users\bkara\Projects\";

if(result.Succeeded)
{
    using var fs = new FileStream(Path.Combine(path, "QuickXL.xlsx"), FileMode.CreateNew, FileAccess.ReadWrite);

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

