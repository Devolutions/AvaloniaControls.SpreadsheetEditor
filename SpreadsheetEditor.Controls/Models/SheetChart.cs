namespace SpreadsheetEditor.Controls.Models;

using System;
using System.Collections.Generic;
using System.Globalization;

public enum ChartKind { Line, Spline, StepLine, Area, StackedArea, Bar, StackedBar, Scatter, Pie, Donut, Radar }

// How the value (Y) axis and data labels format their numbers.
public enum AxisNumberFormat { Number, Currency, Percent }

// A rectangular cell range (0-based, inclusive spans).
public readonly record struct CellRange(int Row, int Column, int RowSpan, int ColumnSpan)
{
    // The A1-style label, e.g. "B6:G10" (or "B6" for a single cell).
    public string ToA1()
    {
        string topLeft = $"{CellAddress.ColumnToLetters(this.Column)}{this.Row + 1}";
        if (this.RowSpan <= 1 && this.ColumnSpan <= 1)
        {
            return topLeft;
        }

        return $"{topLeft}:{CellAddress.ColumnToLetters(this.Column + this.ColumnSpan - 1)}{this.Row + this.RowSpan}";
    }

    // Parses an A1 range like "B6:G10" or a single cell "B6" (case-insensitive, $ ignored).
    public static bool TryParse(string? text, out CellRange range)
    {
        range = default;
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        string[] parts = text.Replace("$", string.Empty).Trim().Split(':');
        if (!TryCell(parts[0], out int r1, out int c1))
        {
            return false;
        }

        (int r2, int c2) = (r1, c1);
        if (parts.Length > 1 && !TryCell(parts[1], out r2, out c2))
        {
            return false;
        }

        range = new CellRange(Math.Min(r1, r2), Math.Min(c1, c2), Math.Abs(r2 - r1) + 1, Math.Abs(c2 - c1) + 1);
        return true;
    }

    private static bool TryCell(string s, out int row, out int column)
    {
        row = column = 0;
        s = s.Trim();
        int i = 0;
        while (i < s.Length && char.IsLetter(s[i]))
        {
            i++;
        }

        if (i == 0 || i == s.Length)
        {
            return false;
        }

        if (!int.TryParse(s[i..], NumberStyles.None, CultureInfo.InvariantCulture, out int rowNumber) || rowNumber < 1)
        {
            return false;
        }

        try
        {
            column = CellAddress.LettersToColumn(s[..i]);
        }
        catch
        {
            return false;
        }

        row = rowNumber - 1;
        return true;
    }
}

// One chart series: a display title and the cell range holding its values. CachedValues, when present,
// is the snapshot the file saved with the chart (used when the referenced cells aren't part of the
// loaded sheet, e.g. an off-screen data table).
public sealed record ChartSeriesSpec(string Title, CellRange Values, IReadOnlyList<double>? CachedValues = null);

// A chart embedded in a sheet. Carries explicit series (each a value range) plus an optional shared
// category range — matching how Excel charts reference data — and a floating placement (anchor cell
// + pixel offset + size). The grid reads the live cell values in those ranges to build/update it.
public sealed class SheetChart
{
    public required ChartKind Kind { get; init; }

    public required IReadOnlyList<ChartSeriesSpec> Series { get; init; }

    public CellRange? Categories { get; init; }

    // Optional chart title shown above the plot.
    public string? Title { get; init; }

    // Whether the chart shows a legend.
    public bool ShowLegend { get; init; } = true;

    // Whether the chart shows value data labels on points/bars.
    public bool ShowDataLabels { get; init; }

    // Number format for the value (Y) axis labels + data labels. Default is a plain number
    // (Excel-style charts aren't currency unless the data is).
    public AxisNumberFormat ValueAxisFormat { get; init; } = AxisNumberFormat.Number;

    // Optional fixed value-axis scale. Null = auto (the chart picks the bounds/step).
    public double? AxisMinimum { get; init; }

    public double? AxisMaximum { get; init; }

    public double? AxisMajorStep { get; init; }

    // The builder inputs (for re-editing via the chart's settings button); null for imported charts.
    public CellRange? SourceRange { get; init; }

    public bool FirstRowHeaders { get; init; }

    public bool FirstColCategories { get; init; }

    // Category labels saved with the chart (preferred over Categories for imported charts whose
    // category cells aren't part of the loaded sheet).
    public IReadOnlyList<string>? CachedCategories { get; init; }

    // Placement: anchored to a from-cell (+ pixel offset). When the to-cell is set, the grid sizes the
    // chart to span from→to in ITS layout (so it covers the same cells as Excel); otherwise Width/Height.
    // Settable so the chart can be dragged/repositioned after creation.
    public required int AnchorRow { get; set; }

    public required int AnchorColumn { get; set; }

    public double OffsetX { get; set; }

    public double OffsetY { get; set; }

    public int? ToRow { get; set; }

    public int? ToColumn { get; set; }

    public double ToOffsetX { get; set; }

    public double ToOffsetY { get; set; }

    public double Width { get; set; } = 380;

    public double Height { get; set; } = 260;

    // Repositions the chart so its top-left corner snaps flush into the given cell, keeping its current
    // pixel size. Clears the from→to span (pins the explicit size) and zeroes the cell offset.
    public void MoveTo(int anchorRow, int anchorColumn, double width, double height)
    {
        this.AnchorRow = Math.Max(0, anchorRow);
        this.AnchorColumn = Math.Max(0, anchorColumn);
        this.OffsetX = 0;
        this.OffsetY = 0;
        this.ToRow = null;
        this.ToColumn = null;
        this.ToOffsetX = 0;
        this.ToOffsetY = 0;
        this.Width = width;
        this.Height = height;
    }

    // Resizes the chart to an explicit pixel size, keeping its anchor. Clears the from→to span (which
    // would otherwise re-derive the size from cells) so the explicit size sticks.
    public void ResizeTo(double width, double height)
    {
        this.ToRow = null;
        this.ToColumn = null;
        this.ToOffsetX = 0;
        this.ToOffsetY = 0;
        this.Width = width;
        this.Height = height;
    }
}
