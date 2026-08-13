namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ClosedXML.Excel;

// Translates Excel Table structured references (e.g. Income[[#This Row],[Jan]:[Dec]], Expenses[[#Totals],
// [Total]], or a bare [Jan]) into plain A1 ranges, so the (structured-reference-unaware) formula engine
// can evaluate them. Done at import time using the workbook's table definitions plus each formula cell's
// position (for [#This Row] and implicit-table references). Only same-sheet tables are handled.
internal static class StructuredReferences
{
    // name[ body ] — the name (table) is optional (absent → the table containing the formula cell). The
    // body is either a single column (Jan), a column range ([Jan]:[Dec]), and/or a [#Section] selector.
    private static readonly Regex Pattern = new(
        @"(?<name>[A-Za-z_\\][\w.]*)?\[(?<body>(?:\[[^\]]*\]|[^\[\]])*)\]",
        RegexOptions.Compiled);

    private static readonly Regex BracketContent = new(@"\[([^\]]+)\]", RegexOptions.Compiled);

    internal sealed record TableInfo(
        string Name, int HeaderRow, int FirstDataRow, int LastDataRow, int TotalsRow,
        int FirstRow, int LastRow, int FirstColumn, int LastColumn, Dictionary<string, int> Columns)
    {
        public bool Contains(int row, int column) =>
            row >= this.FirstRow && row <= this.LastRow && column >= this.FirstColumn && column <= this.LastColumn;
    }

    // Reads the worksheet's table definitions (name, row bands, and column-name → absolute-column map).
    public static List<TableInfo> ReadTables(IXLWorksheet? sheet)
    {
        List<TableInfo> tables = new();
        if (sheet is null)
        {
            return tables;
        }

        foreach (IXLTable table in sheet.Tables)
        {
            try
            {
                IXLRangeAddress range = table.RangeAddress;
                int headerRow = table.ShowHeaderRow ? table.HeadersRow().RowNumber() : range.FirstAddress.RowNumber;
                int totalsRow = table.ShowTotalsRow ? table.TotalsRow().RowNumber() : -1;
                IXLRange? data = table.DataRange;
                int firstData = data?.FirstRow().RowNumber() ?? headerRow + 1;
                int lastData = data?.LastRow().RowNumber() ?? range.LastAddress.RowNumber;

                Dictionary<string, int> columns = new(StringComparer.OrdinalIgnoreCase);
                foreach (IXLTableField field in table.Fields)
                {
                    columns[field.Name] = field.Column.ColumnNumber();
                }

                tables.Add(new TableInfo(
                    table.Name, headerRow, firstData, lastData, totalsRow,
                    range.FirstAddress.RowNumber, range.LastAddress.RowNumber,
                    range.FirstAddress.ColumnNumber, range.LastAddress.ColumnNumber, columns));
            }
            catch
            {
                // Skip a table we can't interpret rather than failing the whole open.
            }
        }

        return tables;
    }

    // Rewrites a formula's structured references to A1. Untranslatable references are left unchanged.
    public static string Translate(string formula, int cellRow, int cellColumn, IReadOnlyList<TableInfo> tables)
    {
        if (string.IsNullOrEmpty(formula) || tables.Count == 0 || !formula.Contains('['))
        {
            return formula;
        }

        return Pattern.Replace(formula, match =>
        {
            string name = match.Groups["name"].Value;
            string body = match.Groups["body"].Value;

            TableInfo? table = name.Length > 0
                ? tables.FirstOrDefault(t => t.Name.Equals(name, StringComparison.OrdinalIgnoreCase))
                : tables.FirstOrDefault(t => t.Contains(cellRow, cellColumn));
            if (table is null)
            {
                return match.Value;
            }

            string section = "data";
            List<string> columnNames = new();
            foreach (string part in body.Split(','))
            {
                MatchCollection brackets = BracketContent.Matches(part);
                if (brackets.Count == 0)
                {
                    string bare = part.Trim();
                    if (bare.Length > 0)
                    {
                        columnNames.Add(bare);
                    }

                    continue;
                }

                foreach (Match bracket in brackets)
                {
                    string token = bracket.Groups[1].Value.Trim();
                    if (token.StartsWith('#'))
                    {
                        section = token[1..].Replace(" ", string.Empty).ToLowerInvariant();
                    }
                    else
                    {
                        columnNames.Add(token);
                    }
                }
            }

            if (columnNames.Count == 0
                || !table.Columns.TryGetValue(columnNames[0], out int firstColumn)
                || !table.Columns.TryGetValue(columnNames[^1], out int lastColumn))
            {
                return match.Value;
            }

            (int rowStart, int rowEnd) = section switch
            {
                "thisrow" => (cellRow, cellRow),
                "totals" => (table.TotalsRow, table.TotalsRow),
                "headers" => (table.HeaderRow, table.HeaderRow),
                "all" => (table.HeaderRow, table.TotalsRow > 0 ? table.TotalsRow : table.LastDataRow),
                _ => (table.FirstDataRow, table.LastDataRow),
            };

            if (rowStart <= 0 || rowEnd <= 0)
            {
                return match.Value;
            }

            string topLeft = $"{ColumnLetters(firstColumn)}{rowStart}";
            string bottomRight = $"{ColumnLetters(lastColumn)}{rowEnd}";
            return topLeft == bottomRight ? topLeft : $"{topLeft}:{bottomRight}";
        });
    }

    // 1-based column number → A1 letters.
    private static string ColumnLetters(int column)
    {
        string result = string.Empty;
        while (column > 0)
        {
            int rem = (column - 1) % 26;
            result = (char)('A' + rem) + result;
            column = (column - 1) / 26;
        }

        return result;
    }
}
