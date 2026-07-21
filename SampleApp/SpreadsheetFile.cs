namespace SampleApp;

using System;
using System.Collections.Generic;
using System.Text;

// App-side open/save for the sample host. Value-only TSV/CSV so the demo can round-trip a sheet.
// Lives in the app (not the control library) because file I/O is a host concern in v1.
public static class SpreadsheetFile
{
    public static char DelimiterFor(string fileName) =>
        fileName.EndsWith(".csv", StringComparison.OrdinalIgnoreCase) ? ',' : '\t';

    // Parses delimited text into rows of cell strings (no quote handling for v1 TSV).
    public static List<List<string>> Parse(string text, char delimiter)
    {
        List<List<string>> rows = new();
        string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
        foreach (string line in normalized.Split('\n'))
        {
            rows.Add(new List<string>(line.Split(delimiter)));
        }

        // drop a trailing empty line produced by a final newline
        if (rows.Count > 1 && rows[^1].Count == 1 && rows[^1][0].Length == 0)
        {
            rows.RemoveAt(rows.Count - 1);
        }

        return rows;
    }

    public static string ToDelimited(IReadOnlyList<IReadOnlyList<string>> rows, char delimiter)
    {
        StringBuilder sb = new();
        foreach (IReadOnlyList<string> row in rows)
        {
            sb.Append(string.Join(delimiter, row)).Append('\n');
        }

        return sb.ToString();
    }
}
