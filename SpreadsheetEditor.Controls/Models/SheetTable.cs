namespace SpreadsheetEditor.Controls.Models;

using System.Collections.Generic;

// Native Excel tables are independent regions, not worksheet-wide filters. Keep their definitions
// through previews, lazy loading and saves so dropdowns do not disappear after reopening a file.
public sealed record SheetTable
{
    public required string Name { get; init; }
    public required CellRange Range { get; init; }
    public IReadOnlyList<string> ColumnNames { get; init; } = new List<string>();
    public CellRange? AutoFilterRange { get; init; }
    public bool ShowHeaderRow { get; init; } = true;
    public bool ShowTotalsRow { get; init; }
    public string? StyleName { get; init; }
    public bool ShowRowStripes { get; init; }
    public bool ShowColumnStripes { get; init; }
    public bool EmphasizeFirstColumn { get; init; }
    public bool EmphasizeLastColumn { get; init; }
}

// Absolute worksheet coordinates; different tables may filter the same column independently.
public sealed record SheetColumnFilter(int HeaderRow, int Column, IReadOnlyList<string> Values);
