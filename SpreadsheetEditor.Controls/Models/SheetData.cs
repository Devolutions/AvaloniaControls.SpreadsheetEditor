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

    // Optional worksheet tab color from the workbook theme/direct color.
    public Color? TabColor { get; init; }
}