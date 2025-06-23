# QuickXL

QuickXL is a lightweight, zero-interop .NET library for generating and consuming Excel (`.xlsx`) files via the Open XML SDK. It exposes a simple, fluent API so you can:

- Export POCO collections to styled `.xlsx` in-memory  
- Import worksheet rows back into `List<T>` via attribute-decorated POCOs  
- Automatically calculate column widths and apply basic styles  
- Work entirely in .NET (no COM/Office interop)  

## Supported Frameworks

- .NET 8.0
- .NET Standard 2.0 (includes .NET Core 3.1, 5, 6, 7, .NET Framework 4.6.1+, Xamarin, Mono…)

## Installation

Install the package from NuGet:

```bash
dotnet add package QuickXL
```

## Quick Start: Export

```csharp
using QuickX.Exporter;

// 1) Create export builder, optionally configure sheet name
// 2) Supply your data
// 3) Define columns by selector (no reflection in loop)
// 4) Call Export() to get a byte[] ready to stream or save
byte[] excel = XLExporter
    .Create<Person>(cfg => cfg.SheetName = "Report")
    .WithData(data)
    .AddColumn(x => x.Name)
    .AddColumn(x => x.Age)
    .AddColumn(x => x.IsActive, opts => opts.HeaderName = "Active?")
    .Export();

// Save to disk:
File.WriteAllBytes("report.xlsx", excel);

```

## Quick Start: Import

```csharp
using QuickX.Importer;

// Open your Excel stream
using var fs = File.OpenRead("report.xlsx");

// 1) Create import builder, optionally configure header row or sheet
// 2) Specify source via FromStream or FromFile
// 3) Call Import() to get List<Person>
var list = XLImporter
    .Create<Person>(cfg => cfg.HeaderRowSettings.StartsAt = 0)
    .FromStream(fs)
    .Import();

```

## Advanced: Export Configuration

You can customize styles globally or per-column via `XLGeneralStyle` and `ColumnSettings`:

```csharp
var bytes = XLExporter
    .Create<MyDto>()
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

