using NPOI.SS.UserModel;
using QuickXL;
using QuickXL.Core.Models.Colors;
using QuickXL.Importer.Examples;


var data = PopulateData();


var exporter = new XLExport<Employee>()
    .CreateBuilder()
    .WithData(data)
    .AddColumn(x => x.Name, cfg =>
    {
        cfg.AutoSizeColumns = true;
        cfg.HeaderName = "Test";
        cfg.AllowEmptyCells = false;
        cfg.HeaderStyle = new()
        {
            Bold = true,
            BackgroundColor = XLColor.Make(255, 0, 0),
            ForegroundColor = XLColor.Make(255, 255, 0),
            Italic = true,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            FontName = "Bauhaus 93",
            FontSize = 25,
            Strikeout = true,
            FillPattern = FillPattern.Diamonds
        };
        cfg.CellStyle = new()
        {            
            Bold = true,
            BackgroundColor = XLColor.Make(255, 0, 255),
            ForegroundColor = XLColor.Make(255, 255, 0),
            Italic = true,
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
            FontName = "Bauhaus 93",
            FontSize = 25,
            Strikeout = true,
            FillPattern = FillPattern.ThickHorizontalBands
        };
    })
    .AddColumn(x => x.Surname, cfg =>
    {
        cfg.AllowEmptyCells = false;
    })
    .AddColumn(x => x.Age, cfg => 
    { 
        cfg.AllowEmptyCells = false; 
    })
    .Build(cfg =>
    {
        cfg.SheetName = "Test_1234";
        cfg.FirstRowIndex = 0;
    });


var result = exporter.Export();

string path = @"C:\Users\TeknikMedya\Projects";

if(result.Succeeded)
{
    using var fs = new FileStream(Path.Combine(path, "QuickXL.xlsx"), FileMode.Create, FileAccess.ReadWrite);

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

