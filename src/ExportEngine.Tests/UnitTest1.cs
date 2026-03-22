using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace ExportEngine.Tests
{
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Stock { get; set; }
        public DateTime CreatedDate { get; set; }
    }

    public class CsvExportTests
    {
        private List<Product> GetTestData() => new List<Product>
        {
            new Product { Id = 1, Name = "Widget", Price = 10.50m, Stock = 100, CreatedDate = new DateTime(2024, 1, 15) },
            new Product { Id = 2, Name = "Gadget", Price = 25.99m, Stock = 50, CreatedDate = new DateTime(2024, 2, 20) },
            new Product { Id = 3, Name = "Doohickey", Price = 5.00m, Stock = 200, CreatedDate = new DateTime(2024, 3, 10) },
        };

        [Fact]
        public void ToCsvString_AutoDetect_IncludesAllColumns()
        {
            var csv = Export.From(GetTestData()).ToCsvString();

            Assert.Contains("Id", csv);
            Assert.Contains("Name", csv);
            Assert.Contains("Price", csv);
            Assert.Contains("Widget", csv);
        }

        [Fact]
        public void ToCsvString_CustomColumns_OnlyIncludesSpecified()
        {
            var csv = Export.From(GetTestData())
                .Column("Product", x => x.Name)
                .Column("Amount", x => x.Price)
                .ToCsvString();

            Assert.Contains("Product", csv);
            Assert.Contains("Amount", csv);
            Assert.DoesNotContain("Stock", csv);
        }

        [Fact]
        public void ToCsvString_WithFormat_AppliesFormat()
        {
            var csv = Export.From(GetTestData())
                .Column("Price", x => x.Price, "F2")
                .ToCsvString();

            Assert.Contains("10.50", csv);
        }

        [Fact]
        public void ToCsvString_EscapesCommas()
        {
            var data = new List<Product>
            {
                new Product { Id = 1, Name = "Widget, Premium", Price = 10.50m }
            };

            var csv = Export.From(data)
                .Column("Name", x => x.Name)
                .ToCsvString();

            Assert.Contains("\"Widget, Premium\"", csv);
        }

        [Fact]
        public void ToCsv_WritesToFile()
        {
            var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.csv");
            try
            {
                Export.From(GetTestData()).ToCsv(path);
                Assert.True(File.Exists(path));
                var content = File.ReadAllText(path);
                Assert.Contains("Widget", content);
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void ToCsvString_SemicolonDelimiter()
        {
            var csv = Export.From(GetTestData())
                .Column("Name", x => x.Name)
                .Column("Price", x => x.Price)
                .ToCsvString(";");

            var lines = csv.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
            Assert.Contains(";", lines[0]); // header has semicolon
        }

        [Fact]
        public void ToCsvString_EmptyData_OnlyHeaders()
        {
            var csv = Export.From(new List<Product>())
                .Column("Name", x => x.Name)
                .ToCsvString();

            var lines = csv.Trim().Split('\n');
            Assert.Single(lines); // only header
        }
    }

    public class ExcelExportTests
    {
        private List<Product> GetTestData() => new List<Product>
        {
            new Product { Id = 1, Name = "Widget", Price = 10.50m, Stock = 100, CreatedDate = new DateTime(2024, 1, 15) },
            new Product { Id = 2, Name = "Gadget", Price = 25.99m, Stock = 50, CreatedDate = new DateTime(2024, 2, 20) },
        };

        [Fact]
        public void ToExcel_CreatesFile()
        {
            var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            try
            {
                Export.From(GetTestData())
                    .SheetName("Products")
                    .Title("Product Report")
                    .ToExcel(path);

                Assert.True(File.Exists(path));
                Assert.True(new FileInfo(path).Length > 0);
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void ToExcelBytes_ReturnsNonEmpty()
        {
            var bytes = Export.From(GetTestData())
                .Column("Name", x => x.Name)
                .Column("Price", x => x.Price, "#,##0.00")
                .ToExcelBytes();

            Assert.NotEmpty(bytes);
            Assert.True(bytes.Length > 100); // valid xlsx
        }

        [Fact]
        public void ToExcel_WithSubtitle_CreatesFile()
        {
            var path = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            try
            {
                Export.From(GetTestData())
                    .Title("Monthly Report")
                    .Subtitle("January 2024")
                    .ToExcel(path);

                Assert.True(File.Exists(path));
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void ToExcel_AutoDetectColumns()
        {
            var bytes = Export.From(GetTestData()).ToExcelBytes();
            Assert.NotEmpty(bytes);
        }
    }

    public class MultiSheetTests
    {
        [Fact]
        public void MultiSheet_CreatesFile()
        {
            var products = new List<Product>
            {
                new Product { Id = 1, Name = "Widget", Price = 10.50m },
            };
            var summary = new[] { new { Total = 1, Revenue = 10.50m } };

            var path = Path.Combine(Path.GetTempPath(), $"multi_{Guid.NewGuid()}.xlsx");
            try
            {
                MultiExport.Create()
                    .AddSheet("Products", products, cfg => cfg
                        .Column("Name", x => x.Name)
                        .Column("Price", x => x.Price))
                    .AddSheet("Summary", summary)
                    .SaveAs(path);

                Assert.True(File.Exists(path));
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void MultiSheet_ToBytes_Works()
        {
            var data = new[] { new { X = 1 } };

            var bytes = MultiExport.Create()
                .AddSheet("Data", data)
                .ToBytes();

            Assert.NotEmpty(bytes);
        }
    }
}
