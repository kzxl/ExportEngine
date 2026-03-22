using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ExportEngine
{
    /// <summary>
    /// Excel (.xlsx) export implementation using ClosedXML.
    /// </summary>
    internal static class ExcelExporter
    {
        public static void Export<T>(ExportBuilder<T> builder, string filePath) where T : class
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var wb = CreateWorkbook(builder))
            {
                wb.SaveAs(filePath);
            }
        }

        public static void Export<T>(ExportBuilder<T> builder, Stream stream) where T : class
        {
            using (var wb = CreateWorkbook(builder))
            {
                wb.SaveAs(stream);
            }
        }

        private static XLWorkbook CreateWorkbook<T>(ExportBuilder<T> builder) where T : class
        {
            var wb = new XLWorkbook();
            var ws = wb.AddWorksheet(builder.SheetTitle);

            int currentRow = 1;

            // Title row
            if (!string.IsNullOrEmpty(builder.ReportTitle))
            {
                ws.Cell(currentRow, 1).Value = builder.ReportTitle;
                ws.Cell(currentRow, 1).Style.Font.Bold = true;
                ws.Cell(currentRow, 1).Style.Font.FontSize = 14;
                ws.Range(currentRow, 1, currentRow, builder.Columns.Count)
                    .Merge();
                currentRow++;
            }

            // Subtitle row
            if (!string.IsNullOrEmpty(builder.ReportSubtitle))
            {
                ws.Cell(currentRow, 1).Value = builder.ReportSubtitle;
                ws.Cell(currentRow, 1).Style.Font.Italic = true;
                ws.Cell(currentRow, 1).Style.Font.FontSize = 11;
                ws.Range(currentRow, 1, currentRow, builder.Columns.Count)
                    .Merge();
                currentRow++;
            }

            if (currentRow > 1) currentRow++; // blank row after title

            // Header row
            int headerRow = currentRow;
            for (int i = 0; i < builder.Columns.Count; i++)
            {
                var cell = ws.Cell(headerRow, i + 1);
                cell.Value = builder.Columns[i].Header;
                cell.Style.Font.Bold = true;
                cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#4472C4");
                cell.Style.Font.FontColor = XLColor.White;
                cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                cell.Style.Border.BottomBorder = XLBorderStyleValues.Medium;
            }
            currentRow++;

            // Data rows
            var dataList = builder.Data.ToList();
            for (int row = 0; row < dataList.Count; row++)
            {
                for (int col = 0; col < builder.Columns.Count; col++)
                {
                    var value = builder.Columns[col].Selector(dataList[row]);
                    var cell = ws.Cell(currentRow + row, col + 1);

                    SetCellValue(cell, value);

                    // Apply format
                    if (!string.IsNullOrEmpty(builder.Columns[col].Format))
                    {
                        cell.Style.NumberFormat.Format = builder.Columns[col].Format;
                    }

                    // Zebra striping
                    if (row % 2 == 1)
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#D6E4F0");
                    }
                }
            }

            // Auto-fit columns
            ws.Columns().AdjustToContents();

            // Add auto-filter to header
            if (dataList.Count > 0)
            {
                ws.Range(headerRow, 1, headerRow + dataList.Count, builder.Columns.Count)
                    .SetAutoFilter();
            }

            return wb;
        }

        private static void SetCellValue(IXLCell cell, object value)
        {
            if (value == null)
            {
                cell.Value = "";
                return;
            }

            switch (value)
            {
                case string s:
                    cell.Value = s;
                    break;
                case int i:
                    cell.Value = i;
                    break;
                case long l:
                    cell.Value = l;
                    break;
                case double d:
                    cell.Value = d;
                    break;
                case float f:
                    cell.Value = f;
                    break;
                case decimal dec:
                    cell.Value = (double)dec;
                    break;
                case System.DateTime dt:
                    cell.Value = dt;
                    break;
                case bool b:
                    cell.Value = b ? "Yes" : "No";
                    break;
                default:
                    cell.Value = value.ToString();
                    break;
            }
        }
    }
}
