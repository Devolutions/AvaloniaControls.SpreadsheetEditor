namespace SpreadsheetEditor.Controls.Models;

using System.Collections.Generic;

using Avalonia.Media;

// A whole-sheet snapshot for loading: cell values+formatting, per-column pixel widths, and merged
// regions. Lets a host serializer (e.g. XLSX) hand the control everything needed to render a sheet
// the way the source app did.
public sealed class SheetData
{
    public required IReadOnlyList<IReadOnlyList<CellData>> Cells { get; init; }

    // Pixel width per column (index = column). May be shorter than the column count.
    public IReadOnlyList<double> ColumnWidths { get; init; } = new List<double>();

    // Pixel height per row (index = row). Zero means use the grid's automatic height.
    public IReadOnlyList<double> RowHeights { get; init; } = new List<double>();

    public IReadOnlyList<MergeRange> Merges { get; init; } = new List<MergeRange>();

    // Pictures embedded in the sheet (anchored to cells, sized in pixels).
    public IReadOnlyList<SheetImage> Images { get; init; } = new List<SheetImage>();

    // Charts embedded in the sheet (parsed from the xlsx drawing/chart parts).
    public IReadOnlyList<SheetChart> Charts { get; init; } = new List<SheetChart>();

    // Defined names → their formula text (e.g. NetIncome → "O36", TotalMonthlyIncome → "SUM(...)"),
    // for formulas like =NetIncome.
    public IReadOnlyDictionary<string, string> Names { get; init; } = new Dictionary<string, string>();

    // Conditional-formatting data bars (proportional in-cell bars).
    public IReadOnlyList<SheetDataBar> DataBars { get; init; } = new List<SheetDataBar>();

    // Numeric conditional font rules, evaluated against current cell values when rendered.
    public IReadOnlyList<SheetCellValueRule> CellValueRules { get; init; } = new List<SheetCellValueRule>();

    // Worksheet AutoFilter range (including its header row), if present.
    public CellRange? AutoFilterRange { get; init; }

    public IReadOnlyList<SheetTable> Tables { get; init; } = new List<SheetTable>();

    public IReadOnlyList<SheetColumnFilter> ColumnFilters { get; init; } = new List<SheetColumnFilter>();

    // Zero-based row indices hidden by filtering (or hidden in the source workbook).
    public IReadOnlyList<int> HiddenRows { get; init; } = new List<int>();

    // Optional worksheet tab color from the workbook theme/direct color.
    public Color? TabColor { get; init; }
}
