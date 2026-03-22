using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExportEngine
{
    /// <summary>
    /// Fluent builder for exporting data to various formats (CSV, Excel).
    /// </summary>
    /// <example>
    /// <code>
    /// Export.From(products)
    ///     .Column("Name", x => x.Name)
    ///     .Column("Price", x => x.Price, format: "#,##0.00")
    ///     .Column("Stock", x => x.Stock)
    ///     .SheetName("Products")
    ///     .Title("Product Report")
    ///     .ToExcel("report.xlsx");
    /// </code>
    /// </example>
    public class ExportBuilder<T> where T : class
    {
        internal IEnumerable<T> Data { get; }
        internal List<ColumnDefinition<T>> Columns { get; } = new List<ColumnDefinition<T>>();
        internal string SheetTitle { get; set; } = "Sheet1";
        internal string ReportTitle { get; set; }
        internal string ReportSubtitle { get; set; }
        internal bool AutoDetectColumns { get; set; } = true;

        internal ExportBuilder(IEnumerable<T> data)
        {
            Data = data ?? throw new ArgumentNullException(nameof(data));
        }

        /// <summary>
        /// Adds a column with a custom value selector.
        /// </summary>
        public ExportBuilder<T> Column(string header, Func<T, object> selector, string format = null)
        {
            AutoDetectColumns = false;
            Columns.Add(new ColumnDefinition<T>
            {
                Header = header,
                Selector = selector,
                Format = format
            });
            return this;
        }

        /// <summary>
        /// Sets the Excel sheet name.
        /// </summary>
        public ExportBuilder<T> SheetName(string name)
        {
            SheetTitle = name;
            return this;
        }

        /// <summary>
        /// Sets a title row at the top of the sheet.
        /// </summary>
        public ExportBuilder<T> Title(string title)
        {
            ReportTitle = title;
            return this;
        }

        /// <summary>
        /// Sets a subtitle row below the title.
        /// </summary>
        public ExportBuilder<T> Subtitle(string subtitle)
        {
            ReportSubtitle = subtitle;
            return this;
        }

        /// <summary>
        /// Exports data to a CSV file.
        /// </summary>
        public void ToCsv(string filePath, string delimiter = ",")
        {
            EnsureColumns();
            CsvExporter.Export(this, filePath, delimiter);
        }

        /// <summary>
        /// Exports data to a CSV stream.
        /// </summary>
        public void ToCsv(Stream stream, string delimiter = ",")
        {
            EnsureColumns();
            CsvExporter.Export(this, stream, delimiter);
        }

        /// <summary>
        /// Exports data to an Excel file (.xlsx).
        /// </summary>
        public void ToExcel(string filePath)
        {
            EnsureColumns();
            ExcelExporter.Export(this, filePath);
        }

        /// <summary>
        /// Exports data to an Excel stream (.xlsx).
        /// </summary>
        public void ToExcel(Stream stream)
        {
            EnsureColumns();
            ExcelExporter.Export(this, stream);
        }

        /// <summary>
        /// Returns the CSV content as a string.
        /// </summary>
        public string ToCsvString(string delimiter = ",")
        {
            EnsureColumns();
            using (var ms = new MemoryStream())
            {
                CsvExporter.Export(this, ms, delimiter);
                ms.Position = 0;
                using (var reader = new StreamReader(ms))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        /// <summary>
        /// Returns the Excel content as a byte array.
        /// </summary>
        public byte[] ToExcelBytes()
        {
            EnsureColumns();
            using (var ms = new MemoryStream())
            {
                ExcelExporter.Export(this, ms);
                return ms.ToArray();
            }
        }

        // ─── Internal ─────────────────────────────────────────────

        internal void EnsureColumns()
        {
            if (Columns.Count > 0) return;
            if (!AutoDetectColumns) return;

            // Auto-detect from public properties
            var props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .Where(p => p.CanRead && IsSimpleType(p.PropertyType));

            foreach (var prop in props)
            {
                var p = prop; // closure capture
                Columns.Add(new ColumnDefinition<T>
                {
                    Header = prop.Name,
                    Selector = obj => p.GetValue(obj)
                });
            }
        }

        private static bool IsSimpleType(Type type)
        {
            type = Nullable.GetUnderlyingType(type) ?? type;
            return type.IsPrimitive || type.IsEnum
                || type == typeof(string) || type == typeof(decimal)
                || type == typeof(DateTime) || type == typeof(DateTimeOffset)
                || type == typeof(Guid);
        }
    }

    /// <summary>
    /// Column definition for export.
    /// </summary>
    public class ColumnDefinition<T>
    {
        /// <summary>Column header text.</summary>
        public string Header { get; set; }

        /// <summary>Value selector function.</summary>
        public Func<T, object> Selector { get; set; }

        /// <summary>Number/date format string (Excel only).</summary>
        public string Format { get; set; }
    }

    /// <summary>
    /// Entry point for the fluent export API.
    /// </summary>
    public static class Export
    {
        /// <summary>
        /// Creates an export builder from a data collection.
        /// Auto-detects columns from public properties if none are specified.
        /// </summary>
        public static ExportBuilder<T> From<T>(IEnumerable<T> data) where T : class
        {
            return new ExportBuilder<T>(data);
        }
    }
}
