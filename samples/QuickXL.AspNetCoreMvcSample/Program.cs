using QuickXL;
using QuickXL.AspNetCoreMvcSample;

var builder = WebApplication.CreateBuilder(args);

// 1) Register QuickXL in the DI container with default settings
builder.Services.AddQuickXL(options =>
{
    options.DefaultSheetName = "Sheet1";
    options.DefaultFirstRowIndex = 0;
});

var app = builder.Build();

// 2) Define a minimal GET endpoint at "/export" that returns an Excel file
app.MapGet("/export", (IXLExporter exporter) =>
{
    // Sample data
    List<Person> people =
    [
        new() { Name = "Alice", Age = 30 },
        new() { Name = "Bob",   Age = 25 }
    ];

    // Build the Excel file
    var excelBytes = exporter
        .CreateBuilder<Person>()
        .WithData(people)
        .AddColumn(p => p.Name)
        .AddColumn(p => p.Age)
        .Export();

    // Send the file back to the client
    return Results.File(
        excelBytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Report.xlsx"
    );
});

app.Run();