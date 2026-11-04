namespace SpreadsheetEditor.Controls.Models;

using System.Collections.Generic;

// A chart resolved into the pieces the .xlsx chart writer needs: its kind/title/placement plus each
// series' value-range formula + cached numbers and the shared category formula + labels. Produced by
// the grid (which resolves the live cell values) and consumed by XlsxChartWriter.
public sealed record ChartWriteData(
    ChartKind Kind,
    string? Title,
    bool ShowLegend,
    bool ShowDataLabels,
    double? AxisMinimum,
    int AnchorRow,
    int AnchorColumn,
    double OffsetXPx,
    double OffsetYPx,
    double WidthPx,
    double HeightPx,
    IReadOnlyList<ChartWriteSeries> Series,
    string? CategoryFormula,
    IReadOnlyList<string> CategoryLabels)
{
    internal string? SourceXml { get; init; }
    public bool HorizontalBars { get; init; }
    public string? ValueAxisTitle { get; init; }
    public string? SecondaryAxisTitle { get; init; }
    public string? CategoryAxisTitle { get; init; }
}

public sealed record ChartWriteSeries(string Title, string ValueFormula, IReadOnlyList<double> Values);
