namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Charts;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Models;

// One chart series: a title and its numeric values (aligned to the shared category list).
public sealed record ChartSeriesData(string Title, IReadOnlyList<double> Values);

// Builds an Avalonia Pro chart control from extracted spreadsheet data. Isolated here (and wrapped in
// try/catch) so the Pro Avalonia.Controls.Charts dependency — which, like the RichTextEditor,
// can't initialize under headless and needs a license — can't crash the grid; returns null on failure.
public static class ChartFactory
{
    // Office-like series colors. Series (and pie slices) must be given an explicit brush — otherwise
    // they render with a null brush and the chart looks empty.
    private static readonly Avalonia.Media.Color[] Palette =
    {
        Avalonia.Media.Color.FromArgb(0xFF, 0x44, 0x72, 0xC4), Avalonia.Media.Color.FromArgb(0xFF, 0xED, 0x7D, 0x31),
        Avalonia.Media.Color.FromArgb(0xFF, 0xA5, 0xA5, 0xA5), Avalonia.Media.Color.FromArgb(0xFF, 0xFF, 0xC0, 0x00),
        Avalonia.Media.Color.FromArgb(0xFF, 0x5B, 0x9B, 0xD5), Avalonia.Media.Color.FromArgb(0xFF, 0x70, 0xAD, 0x47),
        Avalonia.Media.Color.FromArgb(0xFF, 0x26, 0x44, 0x78), Avalonia.Media.Color.FromArgb(0xFF, 0x9E, 0x48, 0x0E),
        Avalonia.Media.Color.FromArgb(0xFF, 0x63, 0x63, 0x63), Avalonia.Media.Color.FromArgb(0xFF, 0x99, 0x73, 0x00),
    };

    private static IBrush PaletteBrush(int index) => new SolidColorBrush(Palette[index % Palette.Length]);

    // A bindable data point (public props so the chart's ValuePath/CategoryPath reflection works).
    public sealed class ChartPoint
    {
        public string Category { get; set; } = string.Empty;

        public double Value { get; set; }

        public IBrush? Brush { get; set; }
    }

    // The .NET format string for an axis label, e.g. "N0"/"C0"/"P0".
    private static string AxisFormat(AxisNumberFormat f) => f switch
    {
        AxisNumberFormat.Currency => "C0",
        AxisNumberFormat.Percent => "P0",
        _ => "N0",
    };

    public static Control? Build(ChartKind kind, IReadOnlyList<string> categories,
        IReadOnlyList<ChartSeriesData> series, double width, double height,
        bool showDataLabels = false, string? title = null, bool? showLegend = null,
        AxisNumberFormat valueFormat = AxisNumberFormat.Number,
        double? axisMin = null, double? axisMax = null, double? axisStep = null,
        ChartLegendPosition legendPosition = ChartLegendPosition.Bottom)
    {
        if (series.Count == 0)
        {
            return null;
        }

        bool legend = showLegend ?? series.Count > 1;
        try
        {
            Control chart = kind switch
            {
                ChartKind.Pie or ChartKind.Donut => BuildPie(categories, series[0], width, height, donut: kind == ChartKind.Donut, legend),
                ChartKind.Radar => BuildRadar(categories, series, width, height, legend),
                _ => BuildCartesian(kind, categories, series, width, height, showDataLabels, legend, valueFormat, axisMin, axisMax, axisStep),
            };
            if (chart is ChartBase cb)
            {
                if (!string.IsNullOrWhiteSpace(title))
                {
                    cb.Title = title;
                }

                // The sub-builders set ShowLegend; the position must be set explicitly or the control
                // won't place the legend where Excel had it (default isn't a guaranteed visible spot).
                if (legend)
                {
                    cb.LegendPosition = legendPosition switch
                    {
                        ChartLegendPosition.Left => LegendPosition.Left,
                        ChartLegendPosition.Right => LegendPosition.Right,
                        ChartLegendPosition.Top => LegendPosition.Top,
                        ChartLegendPosition.None => LegendPosition.None,
                        _ => LegendPosition.Bottom,
                    };
                }
            }

            return chart;
        }
        catch
        {
            return null;
        }
    }

    private static Control BuildCartesian(ChartKind kind, IReadOnlyList<string> categories,
        IReadOnlyList<ChartSeriesData> series, double width, double height, bool showDataLabels, bool showLegend,
        AxisNumberFormat valueFormat, double? axisMin, double? axisMax, double? axisStep)
    {
        string axisFmt = AxisFormat(valueFormat);
        NumericalAxis valueAxis = new()
        {
            ShowLabels = true,
            ShowGridLines = true,
            LabelFormat = axisFmt,
            Minimum = axisMin,    // null = auto
            Maximum = axisMax,    // null = auto (the chart rounds up, leaving headroom for top labels)
            MajorStep = axisStep,
        };
        // Per-bar value labels: how much horizontal room each bar gets. A value label needs roughly
        // 18px to be legible; when the bars are packed tighter than that no font size can stop them
        // overlapping or clipping, so labels are only shown when they actually fit. The ones that do
        // show are sized to the room available and given side padding so edge labels aren't clipped.
        int totalBars = Math.Max(1, categories.Count * Math.Max(1, series.Count));
        double barSpacing = width / totalBars;
        bool labelsFit = showDataLabels && barSpacing >= 18;
        double labelFontSize = Math.Clamp(barSpacing * 0.3, 6, 12);

        CartesianChart chart = new()
        {
            Width = width,
            Height = height,
            ShowLegend = showLegend,
            // Inset so top/edge value labels have room instead of being clipped at the plot boundary
            // (extra on the sides for the first/last bar's label, a little on top for tall bars).
            Padding = labelsFit ? new Thickness(14, 8, 14, 0) : new Thickness(2, 4, 2, 0),
            // Category labels on X (Income/Expenses, months…); value labels + gridlines on Y. Auto fit
            // lets the chart angle the labels when they'd be too crowded to sit horizontally (Excel-
            // style) instead of silently dropping some; short labels that fit stay horizontal.
            HorizontalAxis = new CategoryAxis
            {
                ShowLabels = true,
                ShowGridLines = false,
                LabelFitMode = LabelFitMode.Auto,
            },
            VerticalAxis = valueAxis,
        };
        for (int i = 0; i < series.Count; i++)
        {
            CartesianSeries cs = MakeCartesianSeries(kind, PaletteBrush(i));
            cs.ItemsSource = Points(categories, series[i].Values);
            cs.CategoryPath = nameof(ChartPoint.Category);
            cs.ValuePath = nameof(ChartPoint.Value);
            cs.Title = series[i].Title;
            if (labelsFit)
            {
                cs.ShowLabels = true;
                cs.LabelFormat = "{0:" + axisFmt + "}"; // placeholder form (a bare "C0" renders literally)
                cs.LabelFontSize = labelFontSize;
                // Dark text on a translucent light pill so labels stay legible whether they sit over a
                // bar or the plot background (no single text color works over both).
                cs.LabelForeground = new SolidColorBrush(Color.FromRgb(0x20, 0x20, 0x20));
                cs.LabelBackground = new SolidColorBrush(Color.FromArgb(0xDC, 0xFF, 0xFF, 0xFF));
                cs.LabelCornerRadius = new CornerRadius(2);
                cs.LabelPadding = new Thickness(2, 0, 2, 0);
            }

            chart.Series.Add(cs);
        }

        return chart;
    }

    private static CartesianSeries MakeCartesianSeries(ChartKind kind, IBrush brush) => kind switch
    {
        ChartKind.Bar => new BarSeries { Fill = brush },
        ChartKind.StackedBar => new StackedBarSeries { Fill = brush },
        ChartKind.Area => new AreaSeries { Fill = Translucent(brush), Stroke = brush, StrokeThickness = 2 },
        ChartKind.StackedArea => new StackedAreaSeries { Fill = Translucent(brush), Stroke = brush, StrokeThickness = 2 },
        ChartKind.Spline => new SplineSeries { Stroke = brush, StrokeThickness = 2, ShowMarkers = false },
        ChartKind.StepLine => new StepLineSeries { Stroke = brush, StrokeThickness = 2, ShowMarkers = false },
        ChartKind.Scatter => new ScatterSeries { MarkerFill = brush, MarkerSize = 6, ShowMarkers = true },
        _ => new LineSeries { Stroke = brush, StrokeThickness = 2, ShowMarkers = false },
    };

    private static Control BuildPie(IReadOnlyList<string> categories, ChartSeriesData series, double width, double height, bool donut, bool showLegend)
    {
        PieChart chart = new() { Width = width, Height = height, ShowLegend = showLegend, InnerRadiusFactor = donut ? 0.5 : 0 };
        List<ChartPoint> points = Points(categories, series.Values);
        for (int i = 0; i < points.Count; i++)
        {
            points[i].Brush = PaletteBrush(i); // colour each slice
        }

        chart.Series.Add(new PieSeries
        {
            ItemsSource = points,
            ValuePath = nameof(ChartPoint.Value),
            LabelPath = nameof(ChartPoint.Category),
            PointBrushPath = nameof(ChartPoint.Brush),
        });
        return chart;
    }

    private static Control BuildRadar(IReadOnlyList<string> categories, IReadOnlyList<ChartSeriesData> series, double width, double height, bool showLegend)
    {
        RadarChart chart = new() { Width = width, Height = height, ShowLegend = showLegend };
        for (int i = 0; i < series.Count; i++)
        {
            IBrush brush = PaletteBrush(i);
            chart.Series.Add(new RadarSeries
            {
                ItemsSource = Points(categories, series[i].Values),
                ValuePath = nameof(ChartPoint.Value),
                Stroke = brush,
                Fill = Translucent(brush),
                Title = series[i].Title,
            });
        }

        return chart;
    }

    private static IBrush Translucent(IBrush brush) =>
        brush is ISolidColorBrush s
            ? new SolidColorBrush(Avalonia.Media.Color.FromArgb(0x55, s.Color.R, s.Color.G, s.Color.B))
            : brush;

    private static List<ChartPoint> Points(IReadOnlyList<string> categories, IReadOnlyList<double> values) =>
        categories.Select((c, i) => new ChartPoint { Category = c, Value = i < values.Count ? values[i] : double.NaN }).ToList();
}
