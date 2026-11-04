namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;

// The chart package's BarSeries is vertical. Keep horizontal Excel bars horizontal without
// rotating the control (which would rotate labels and break pointer/layout coordinates).
internal sealed class HorizontalBarChart(IReadOnlyList<string> categories, IReadOnlyList<ChartSeriesData> series,
    bool stacked, string? title, bool legend) : Control
{
    private static readonly Color[] Palette = { Color.Parse("#4472C4"), Color.Parse("#ED7D31"), Color.Parse("#A5A5A5"), Color.Parse("#FFC000"), Color.Parse("#5B9BD5") };

    public override void Render(DrawingContext context)
    {
        base.Render(context);
        double width = this.Bounds.Width, height = this.Bounds.Height;
        if (width < 100 || height < 80 || categories.Count == 0) return;
        void Text(string text, double x, double y, double size = 10, bool bold = false)
        {
            FormattedText formatted = new(text, CultureInfo.InvariantCulture, FlowDirection.LeftToRight,
                new Typeface(FontFamily.Default, FontStyle.Normal, bold ? FontWeight.Bold : FontWeight.Normal), size, Brushes.Black);
            context.DrawText(formatted, new Point(x, y));
        }
        if (!string.IsNullOrWhiteSpace(title)) Text(title, 12, 3, 12, true);
        double top = string.IsNullOrWhiteSpace(title) ? 8 : 25;
        double bottom = height - (legend ? 42 : 23), left = Math.Min(75, width * 0.22), right = width - 12;
        double min = 0, max = 0;
        for (int i = 0; i < categories.Count; i++)
        {
            double[] values = series.Select(s => i < s.Values.Count && double.IsFinite(s.Values[i]) ? s.Values[i] : 0).ToArray();
            min = Math.Min(min, stacked ? values.Where(v => v < 0).Sum() : values.Min());
            max = Math.Max(max, stacked ? values.Where(v => v > 0).Sum() : values.Max());
        }
        if (max <= min) max = min + 1;
        double X(double value) => left + (value - min) / (max - min) * (right - left);
        Pen grid = new(Brushes.LightGray, 0.5);
        for (int tick = 0; tick <= 4; tick++)
        {
            double value = min + (max - min) * tick / 4;
            double x = X(value);
            context.DrawLine(grid, new Point(x, top), new Point(x, bottom));
            Text(value.ToString("0.##", CultureInfo.InvariantCulture), x - 10, bottom + 3, 9);
        }
        double rowHeight = Math.Max(1, (bottom - top) / categories.Count);
        using (context.PushClip(new Rect(0, top, width, Math.Max(1, bottom - top))))
        for (int i = 0; i < categories.Count; i++)
        {
            double y = top + i * rowHeight;
            Text(categories[i], 3, y + Math.Max(0, (rowHeight - 12) / 2), Math.Min(10, rowHeight));
            double positive = 0, negative = 0;
            for (int s = 0; s < series.Count; s++)
            {
                double value = i < series[s].Values.Count && double.IsFinite(series[s].Values[i]) ? series[s].Values[i] : 0;
                double from = stacked ? value >= 0 ? positive : negative : 0;
                double to = from + value;
                if (value >= 0) positive = to; else negative = to;
                double barHeight = rowHeight * 0.72 / (stacked ? 1 : series.Count);
                context.DrawRectangle(new SolidColorBrush(series[s].Color ?? Palette[s % Palette.Length]), null,
                    new Rect(Math.Min(X(from), X(to)), y + rowHeight * 0.14 + (stacked ? 0 : s * barHeight), Math.Abs(X(to) - X(from)), Math.Max(0.5, barHeight)));
            }
        }
        if (legend)
        {
            double x = left;
            for (int s = 0; s < series.Count; s++)
            {
                context.DrawRectangle(new SolidColorBrush(series[s].Color ?? Palette[s % Palette.Length]), null, new Rect(x, height - 14, 9, 9));
                Text(series[s].Title, x + 12, height - 17);
                x += 22 + series[s].Title.Length * 6;
            }
        }
    }
}
