using System;
using System.IO;
using System.Text;

namespace ExportEngine
{
    /// <summary>
    /// High-performance RFC 4180 streaming CSV exporter.
    /// Operates directly on streams without intermediate string allocations.
    /// </summary>
    internal static class CsvExporter
    {
        public static void Export<T>(ExportBuilder<T> builder, string filePath, string delimiter) where T : class
        {
            var dir = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            using (var fs = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 65536))
            {
                Export(builder, fs, delimiter);
            }
        }

        public static void Export<T>(ExportBuilder<T> builder, Stream stream, string delimiter) where T : class
        {
            using (var writer = new StreamWriter(stream, new UTF8Encoding(true), 65536, leaveOpen: true))
            {
                // Header row
                for (int i = 0; i < builder.Columns.Count; i++)
                {
                    if (i > 0) writer.Write(delimiter);
                    WriteEscapedField(writer, builder.Columns[i].Header, delimiter);
                }
                writer.WriteLine();

                // Data rows
                foreach (var item in builder.Data)
                {
                    for (int i = 0; i < builder.Columns.Count; i++)
                    {
                        if (i > 0) writer.Write(delimiter);
                        var value = builder.Columns[i].Selector(item);
                        var str = FormatValue(value, builder.Columns[i].Format);
                        WriteEscapedField(writer, str, delimiter);
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

        private static void WriteEscapedField(TextWriter writer, string value, string delimiter)
        {
            if (string.IsNullOrEmpty(value)) return;

            bool needsQuoting = value.IndexOf(delimiter, StringComparison.Ordinal) >= 0
                || value.IndexOf('"') >= 0
                || value.IndexOf('\n') >= 0
                || value.IndexOf('\r') >= 0;

            if (!needsQuoting)
            {
                writer.Write(value);
                return;
            }

            writer.Write('"');
            int lastIndex = 0;
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] == '"')
                {
                    if (i > lastIndex)
                    {
                        writer.Write(value.Substring(lastIndex, i - lastIndex));
                    }
                    writer.Write("\"\"");
                    lastIndex = i + 1;
                }
            }

            if (lastIndex < value.Length)
            {
                writer.Write(value.Substring(lastIndex));
            }

            writer.Write('"');
        }
    }
}
