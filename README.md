# QuickXL

QuickXL is a lightweight, zero-interop .NET library for generating and consuming Excel (`.xlsx`) files via the Open XML SDK. It exposes a simple, fluent API so you can:

- Define your sheet name and data in code  
- Add columns by property selector (no reflection in the hot loop)  
- Automatically calculate column widths and apply basic styles  
- Export to `byte[]` or stream in a single call  
- Import rows back into `List<T>` via attribute-decorated POCOs  
- 
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

## Quick Start: Export

```csharp
using QuickXL;

// 1) Begin an export session by sheet name
// 2) Supply your data
// 3) Define columns by selector
// 4) Call Export() to get a byte[] ready to write to disk or HTTP response
byte[] excelBytes = ExcelExporter
    .Create<Person>("MySheet")
    .WithData(myPersonList)
    .AddColumn(x => x.Name,     opts => opts.HeaderName = "Full Name")
    .AddColumn(x => x.Age)
    .AddColumn(x => x.IsActive, opts => opts.HeaderName = "Active?")
    .Export();

// Write to file
File.WriteAllBytes("report.xlsx", excelBytes);

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

