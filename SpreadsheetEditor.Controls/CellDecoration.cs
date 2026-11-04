namespace SpreadsheetEditor.Controls;

using System;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Models;

// A single lightweight drawing for patterns and border styles, rather than one control per line.
internal sealed class CellDecoration(CellAppearance appearance, double zoom, bool fill) : Control
{
    public override void Render(DrawingContext context)
    {
        base.Render(context);
        Rect rect = new(this.Bounds.Size);
        if (fill)
        {
            if (appearance.Fill is not { } style) return;
            if (style.Stops.Count > 0)
            {
                double angle = style.Degree * Math.PI / 180;
                LinearGradientBrush gradient = new()
                {
                    StartPoint = new RelativePoint(0.5 - Math.Cos(angle) / 2, 0.5 - Math.Sin(angle) / 2, RelativeUnit.Relative),
                    EndPoint = new RelativePoint(0.5 + Math.Cos(angle) / 2, 0.5 + Math.Sin(angle) / 2, RelativeUnit.Relative),
                };
                foreach (CellGradientStop stop in style.Stops) gradient.GradientStops.Add(new GradientStop(stop.Color, stop.Position));
                context.DrawRectangle(gradient, null, rect);
                return;
            }
            context.DrawRectangle(new SolidColorBrush(style.Background ?? Colors.White), null, rect);
            IBrush brush = new SolidColorBrush(style.Foreground ?? Colors.Gray);
            Pen pen = new(brush, Math.Max(0.5, zoom));
            double step = (style.Pattern is "gray125" ? 8 : style.Pattern is "gray0625" ? 12 : 5) * zoom;
            using (context.PushClip(rect))
            {
                if (style.Pattern is "gray125" or "gray0625")
                {
                    for (double y = 0; y < rect.Height; y += step)
                    {
                        for (double x = 0; x < rect.Width; x += step)
                        {
                            context.DrawRectangle(brush, null, new Rect(x, y, zoom, zoom));
                        }
                    }
                }
                else
                {
                    for (double x = 0; x < rect.Width; x += step) context.DrawLine(pen, new Point(x, 0), new Point(x, rect.Height));
                    for (double y = 0; y < rect.Height; y += step) context.DrawLine(pen, new Point(0, y), new Point(rect.Width, y));
                }
            }
            return;
        }
        void Edge(CellEdge? edge, Point start, Point end, Vector inward)
        {
            if (edge is null) return;
            double width = (edge.Style is "thick" ? 3 : edge.Style.StartsWith("medium", StringComparison.OrdinalIgnoreCase) ? 2 : edge.Style == "hair" ? 0.5 : 1) * zoom;
            IDashStyle? dash = edge.Style switch
            {
                "dashed" or "mediumDashed" => DashStyle.Dash,
                "dotted" => DashStyle.Dot,
                "dashDot" or "mediumDashDot" => DashStyle.DashDot,
                "dashDotDot" or "mediumDashDotDot" => DashStyle.DashDotDot,
                _ => null,
            };
            Pen pen = new(new SolidColorBrush(edge.Color), width, dash);
            Vector offset = inward * (width / 2);
            context.DrawLine(pen, start + offset, end + offset);
            if (edge.Style == "double") context.DrawLine(pen, start + inward * (3 * zoom), end + inward * (3 * zoom));
        }
        Edge(appearance.Left, rect.TopLeft, rect.BottomLeft, new Vector(1, 0));
        Edge(appearance.Right, rect.TopRight, rect.BottomRight, new Vector(-1, 0));
        Edge(appearance.Top, rect.TopLeft, rect.TopRight, new Vector(0, 1));
        Edge(appearance.Bottom, rect.BottomLeft, rect.BottomRight, new Vector(0, -1));
    }

    protected override Type StyleKeyOverride => typeof(Control);
}
