using System;
using System.Collections.Generic;
using System.IO;

namespace ExportEngine
{
    /// <summary>
    /// Multi-sheet Excel export builder.
    /// </summary>
    /// <example>
    /// <code>
    /// MultiExport.Create()
    ///     .AddSheet("Products", products, cfg => cfg
    ///         .Column("Name", x => x.Name)
    ///         .Column("Price", x => x.Price, "#,##0.00"))
    ///     .AddSheet("Summary", summaryData)
    ///     .Title("Monthly Report")
    ///     .SaveAs("report.xlsx");
    /// </code>
    /// </example>
    public class MultiSheetBuilder
    {
        private readonly List<Action<ClosedXML.Excel.XLWorkbook>> _sheetActions
            = new List<Action<ClosedXML.Excel.XLWorkbook>>();

        internal MultiSheetBuilder() { }

        /// <summary>
        /// Adds a sheet with auto-detected columns.
        /// </summary>
        public MultiSheetBuilder AddSheet<T>(string sheetName, IEnumerable<T> data) where T : class
        {
            var builder = Export.From(data).SheetName(sheetName);
            _sheetActions.Add(wb =>
            {
                builder.EnsureColumns();
                AddSheetToWorkbook(wb, builder);
            });
            return this;
        }

        /// <summary>
        /// Adds a sheet with custom column configuration.
        /// </summary>
        public MultiSheetBuilder AddSheet<T>(
            string sheetName,
            IEnumerable<T> data,
            Action<ExportBuilder<T>> configure) where T : class
        {
            var builder = Export.From(data).SheetName(sheetName);
            configure(builder);
            _sheetActions.Add(wb =>
            {
                builder.EnsureColumns();
                AddSheetToWorkbook(wb, builder);
            });
            return this;
        }

        /// <summary>
        /// Saves all sheets to an Excel file.
        /// </summary>
        public void SaveAs(string filePath)
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var wb = new ClosedXML.Excel.XLWorkbook())
            {
                foreach (var action in _sheetActions)
                    action(wb);
                wb.SaveAs(filePath);
            }
        }

        /// <summary>
        /// Returns the Excel content as a byte array.
        /// </summary>
        public byte[] ToBytes()
        {
            using (var ms = new MemoryStream())
            using (var wb = new ClosedXML.Excel.XLWorkbook())
            {
                foreach (var action in _sheetActions)
                    action(wb);
                wb.SaveAs(ms);
                return ms.ToArray();
            }
        }

        private static void AddSheetToWorkbook<T>(
            ClosedXML.Excel.XLWorkbook wb,
            ExportBuilder<T> builder) where T : class
        {
            var ws = wb.AddWorksheet(builder.SheetTitle);

            int currentRow = 1;

            // Header
            for (int i = 0; i < builder.Columns.Count; i++)
            {
                var cell = ws.Cell(currentRow, i + 1);
                cell.Value = builder.Columns[i].Header;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromHtml("#4472C4");
                cell.Style.Font.FontColor = ClosedXML.Excel.XLColor.White;
            }
            currentRow++;

            // Data
            var dataList = new List<T>(builder.Data);
            foreach (var item in dataList)
            {
                for (int col = 0; col < builder.Columns.Count; col++)
                {
                    var value = builder.Columns[col].Selector(item);
                    var cell = ws.Cell(currentRow, col + 1);
                    if (value != null) cell.Value = value.ToString();
                    if (!string.IsNullOrEmpty(builder.Columns[col].Format))
                        cell.Style.NumberFormat.Format = builder.Columns[col].Format;
                }
                currentRow++;
            }

            ws.Columns().AdjustToContents();
        }
    }

    /// <summary>
    /// Entry point for multi-sheet export.
    /// </summary>
    public static class MultiExport
    {
        /// <summary>
        /// Creates a new multi-sheet export builder.
        /// </summary>
        public static MultiSheetBuilder Create()
        {
            return new MultiSheetBuilder();
        }
    }
}
