# QuickXL

QuickXL is a lightweight .NET library for generating Excel (.xlsx) files using the Open XML SDK. It provides a simple, fluent API to define columns, map your data, and export to a `byte[]` or stream, without relying on heavy interop or third‑party dependencies.

## Supported Frameworks

- .NET 8.0
- .NET Standard 2.0 (includes .NET Core 3.1, 5, 6, 7, .NET Framework 4.6.1+, Xamarin, Mono…)

## Features

- Pure .NET Core, no COM or Office installation required
- Zero reflection in export loop: uses precompiled delegates for high performance
- Automatic column width calculation
- Support for styling via OpenXML stylesheet definitions
- Minimal API and ASP.NET Core integration

## Installation

Install the package from NuGet:

```bash
dotnet add package QuickXL
```

> The QuickXL package includes its dependencies (`DocumentFormat.OpenXml`, `Ardalis.GuardClauses`, etc.) transitively.

## ASP.NET Core Minimal API Example

```csharp
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;
using QuickXL;

var builder = WebApplication.CreateBuilder(args);

// Register QuickXL
builder.Services.AddQuickXL(opts =>
{
    opts.DefaultSheetName     = "Sheet1";
    opts.DefaultFirstRowIndex = 0;
});

var app = builder.Build();

app.MapGet("/export", (IXLExporter exporter) =>
{
    Person[] data = new Person[]
    {
        new Person { Name = "Alice", Age = 30 },
        new Person { Name = "Bob",   Age = 25 }
    };

    byte[] bytes = exporter
        .CreateBuilder<Person>()
        .WithData(data)
        .AddColumn("Name", x => x.Name)
        .AddColumn("Age",  x => x.Age)
        .Export();

    return Results.File(
        bytes,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "report.xlsx"
    );
});

app.Run();

public class Person
{
    public string Name { get; set; } = string.Empty;
    public int    Age  { get; set; }
}

```

## Advanced Configuration

You can customize styles globally or per-column via `XLGeneralStyle` and `ColumnSettings`:

```csharp
var bytes = exporter.CreateBuilder<MyDto>()
    .WithData(data)
    .AddGeneralStyle(new XLGeneralStyle
    {
        HeaderStyle = new XLCellStyle { FontSize = 14, Bold = true, ForegroundColor = "#FF0000" },
        CellStyle   = new XLCellStyle { FontName = "Arial" }
    })
    .AddColumn("Value", x => x.Value, opts => opts.HeaderStyle = new XLCellStyle { Italic = true })
    .Export();
```

## Contributing

Contributions are welcome! Feel free to open issues, submit PRs, or suggest new examples under `samples/` or `examples/`.

## License

This project is licensed under the [MIT License](LICENSE).

