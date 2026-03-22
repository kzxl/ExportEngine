# 📊 ExportEngine — Fluent Data Export for .NET

> **Export any collection to CSV or Excel (.xlsx) with a one-liner API.**

[![.NET Standard](https://img.shields.io/badge/.NET%20Standard-2.0-blue)](https://docs.microsoft.com/en-us/dotnet/standard/net-standard)
[![License](https://img.shields.io/badge/License-Apache%202.0-blue.svg)](LICENSE)

ExportEngine provides a **fluent, type-safe API** for exporting data collections to **CSV** and **Excel** formats.
Auto-detects columns from object properties, or configure custom columns with formatting.

---

## 📦 Features

| Feature | Description |
|---------|-------------|
| **Auto Column Detection** | Discovers columns from public properties |
| **Custom Columns** | Select fields with `Column("Header", x => x.Prop)` |
| **Number/Date Formatting** | Apply Excel/string format patterns |
| **Excel Styling** | Blue headers, zebra striping, auto-filter, auto-width |
| **Title & Subtitle** | Add report headers above data |
| **Multi-Sheet** | Multiple datasets in one Excel workbook |
| **CSV Export** | Proper escaping, UTF-8, custom delimiter |
| **Stream Support** | Output to file, stream, or byte array |

---

## 🚀 Quick Start

### Simple CSV Export

```csharp
Export.From(products).ToCsv("products.csv");
```

### Custom Excel Report

```csharp
Export.From(products)
    .Column("Product Name", x => x.Name)
    .Column("Price", x => x.Price, "#,##0.00")
    .Column("In Stock", x => x.Stock)
    .Title("Product Catalog")
    .Subtitle("Generated: " + DateTime.Now.ToShortDateString())
    .SheetName("Products")
    .ToExcel("catalog.xlsx");
```

### Multi-Sheet Workbook

```csharp
MultiExport.Create()
    .AddSheet("Products", products, cfg => cfg
        .Column("Name", x => x.Name)
        .Column("Price", x => x.Price, "#,##0"))
    .AddSheet("Summary", summaryData)
    .SaveAs("report.xlsx");
```

### Return as Bytes (for WebAPI)

```csharp
// Excel
byte[] excelBytes = Export.From(products).ToExcelBytes();
return File(excelBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "report.xlsx");

// CSV
string csvContent = Export.From(products).ToCsvString();
```

---

## 📖 API Reference

### Single Sheet

```csharp
Export.From(data)                           // start builder
    .Column("Header", x => x.Prop, "fmt")  // add column (optional)
    .SheetName("MySheet")                   // Excel sheet name
    .Title("Report Title")                  // title row
    .Subtitle("Subtitle")                   // subtitle row
    .ToCsv("file.csv")                      // export CSV to file
    .ToCsv(stream, ";")                     // export CSV to stream
    .ToCsvString()                          // CSV as string
    .ToExcel("file.xlsx")                   // export Excel to file
    .ToExcel(stream)                        // export Excel to stream
    .ToExcelBytes()                         // Excel as byte[]
```

### Multi-Sheet

```csharp
MultiExport.Create()
    .AddSheet("Name", data)                 // auto columns
    .AddSheet("Name", data, cfg => ...)     // custom columns
    .SaveAs("file.xlsx")                    // save to file
    .ToBytes()                              // as byte[]
```

---

## 📄 License

Apache License 2.0
