using System;
using System.IO;
using System.Text;

namespace ExportEngine
{
    /// <summary>
    /// CSV export implementation.
    /// </summary>
    internal static class CsvExporter
    {
        public static void Export<T>(ExportBuilder<T> builder, string filePath, string delimiter) where T : class
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                Export(builder, fs, delimiter);
            }
        }

        public static void Export<T>(ExportBuilder<T> builder, Stream stream, string delimiter) where T : class
        {
            using (var writer = new StreamWriter(stream, new UTF8Encoding(true), 1024, leaveOpen: true))
            {
                // Header row
                for (int i = 0; i < builder.Columns.Count; i++)
                {
                    if (i > 0) writer.Write(delimiter);
                    writer.Write(EscapeCsv(builder.Columns[i].Header, delimiter));
                }
                writer.WriteLine();

                // Data rows
                foreach (var item in builder.Data)
                {
                    for (int i = 0; i < builder.Columns.Count; i++)
                    {
                        if (i > 0) writer.Write(delimiter);
                        var value = builder.Columns[i].Selector(item);
                        writer.Write(EscapeCsv(FormatValue(value, builder.Columns[i].Format), delimiter));
                    }
                    writer.WriteLine();
                }
            }
        }

        private static string FormatValue(object value, string format)
        {
            if (value == null) return "";
            if (!string.IsNullOrEmpty(format) && value is IFormattable f)
                return f.ToString(format, null);
            return value.ToString();
        }

        private static string EscapeCsv(string value, string delimiter)
        {
            if (string.IsNullOrEmpty(value)) return "";

            bool needsQuoting = value.Contains(delimiter)
                || value.Contains("\"")
                || value.Contains("\n")
                || value.Contains("\r");

            if (!needsQuoting) return value;

            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }
    }
}
