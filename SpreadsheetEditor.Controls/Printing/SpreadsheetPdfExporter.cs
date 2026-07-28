namespace SpreadsheetEditor.Controls.Printing;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Avalonia.Media;
using SkiaSharp;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.ViewModels;

// Renders a sheet's used range to a paginated PDF using SkiaSharp's SKDocument.CreatePdf.
// This is fully cross-platform (Windows/macOS/Linux) and self-contained: it needs no browser
// and no OS print subsystem. Cells are drawn with their fills, borders, alignment, and
// (rich) run formatting; the sheet is scaled down to fit the page width ("fit to 1 page wide")
// and rows flow onto as many pages as needed. Charts and images are not rendered (data only),
// matching the HTML export.
public static class SpreadsheetPdfExporter
{
    // US Letter, portrait, at 72 pt/inch (PDF's native unit).
    private const float PageWidth = 612f;

    private const float PageHeight = 792f;

    private const float Margin = 36f;

    // Screen sizes (column widths, row heights, font sizes) are in device px at 96 dpi;
    // PDF is 72 pt/inch, so 1 px = 0.75 pt.
    private const double PxToPt = 72.0 / 96.0;

    private static readonly SKColor GridLineColor = new(0xC8, 0xC8, 0xC8);

    private static readonly SKColor TextColor = SKColors.Black;

    public static void Export(SpreadsheetViewModel sheet, Stream stream)
    {
        if (sheet is null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        int lastRow = Math.Max(sheet.LastNonEmptyRow(), 0);
        int lastColumn = LastNonEmptyColumn(sheet, lastRow);

        (Dictionary<(int Row, int Column), MergeRange> anchors, HashSet<(int Row, int Column)> covered) = BuildMergeMaps(sheet);

        // Fit-to-width scale: shrink (never grow) so the used columns fit the printable width.
        float contentWidth = PageWidth - (2 * Margin);
        double totalWidthPt = 0;
        for (int c = 0; c <= lastColumn; c++)
        {
            totalWidthPt += ColumnWidthPx(sheet, c) * PxToPt;
        }

        double fitScale = totalWidthPt > contentWidth && totalWidthPt > 0 ? contentWidth / totalWidthPt : 1.0;
        double scale = PxToPt * fitScale;

        // Column x-offsets and widths, in page points, relative to the left margin.
        float[] columnLeft = new float[lastColumn + 2];
        columnLeft[0] = Margin;
        for (int c = 0; c <= lastColumn; c++)
        {
            columnLeft[c + 1] = columnLeft[c] + (float)(ColumnWidthPx(sheet, c) * scale);
        }

        float[] rowHeightPt = new float[lastRow + 1];
        for (int r = 0; r <= lastRow; r++)
        {
            rowHeightPt[r] = (float)(RowHeightPx(sheet, r) * scale);
        }

        const float titleHeight = 24f;
        float gridTop = Margin + titleHeight;
        float contentBottom = PageHeight - Margin;

        Dictionary<TypefaceKey, SKTypeface> typefaces = new();

        try
        {
            SKDocumentPdfMetadata metadata = new() { Title = sheet.Name };
            using SKDocument document = SKDocument.CreatePdf(stream, metadata);

            int row = 0;
            while (row <= lastRow)
            {
                SKCanvas canvas = document.BeginPage(PageWidth, PageHeight);
                try
                {
                    DrawTitle(canvas, sheet.Name);

                    // Pack as many rows as fit under the title on this page (at least one).
                    float y = gridTop;
                    int pageStart = row;
                    while (row <= lastRow && (row == pageStart || y + rowHeightPt[row] <= contentBottom))
                    {
                        y += rowHeightPt[row];
                        row++;
                    }

                    int pageEnd = row; // exclusive

                    // Clip so a merged block spanning past the page bottom is trimmed cleanly.
                    canvas.Save();
                    canvas.ClipRect(new SKRect(Margin, gridTop, PageWidth - Margin, contentBottom));

                    float rowTop = gridTop;
                    for (int r = pageStart; r < pageEnd; r++)
                    {
                        for (int c = 0; c <= lastColumn; c++)
                        {
                            if (covered.Contains((r, c)))
                            {
                                continue;
                            }

                            CellViewModel? cell = sheet.GetCell(new CellAddress(r, c));

                            int columnSpan = 1;
                            int rowSpan = 1;
                            if (anchors.TryGetValue((r, c), out MergeRange span))
                            {
                                columnSpan = span.ColumnSpan;
                                rowSpan = span.RowSpan;
                            }

                            float cellLeft = columnLeft[c];
                            float cellRight = columnLeft[Math.Min(c + columnSpan, lastColumn + 1)];
                            float cellBottom = rowTop;
                            for (int rr = r; rr < r + rowSpan && rr <= lastRow; rr++)
                            {
                                cellBottom += rowHeightPt[rr];
                            }

                            SKRect rect = new(cellLeft, rowTop, cellRight, cellBottom);
                            DrawCell(canvas, cell, rect, (float)scale, typefaces);
                        }

                        rowTop += rowHeightPt[r];
                    }

                    canvas.Restore();
                }
                finally
                {
                    document.EndPage();
                }
            }

            document.Close();
        }
        finally
        {
            foreach (SKTypeface typeface in typefaces.Values)
            {
                typeface.Dispose();
            }
        }
    }

    private static void DrawTitle(SKCanvas canvas, string title)
    {
        using SKTypeface typeface = SKTypeface.FromFamilyName("Calibri", SKFontStyleWeight.Bold, SKFontStyleWidth.Normal, SKFontStyleSlant.Upright);
        using SKFont font = new(typeface, 15f);
        using SKPaint paint = new() { Color = TextColor, IsAntialias = true };
        canvas.DrawText(title ?? string.Empty, Margin, Margin + 15f, font, paint);
    }

    private static void DrawCell(SKCanvas canvas, CellViewModel? cell, SKRect rect, float scale, Dictionary<TypefaceKey, SKTypeface> typefaces)
    {
        // Background fill.
        if (cell?.BackgroundColor is { } bg)
        {
            using SKPaint fill = new() { Color = ToSkColor(bg), Style = SKPaintStyle.Fill };
            canvas.DrawRect(rect, fill);
        }

        // Cell border (thin gray, matching the HTML export's grid lines).
        using (SKPaint border = new() { Color = GridLineColor, Style = SKPaintStyle.Stroke, StrokeWidth = 0.5f, IsAntialias = true })
        {
            canvas.DrawRect(rect, border);
        }

        if (cell is null)
        {
            return;
        }

        IReadOnlyList<CellRun> runs = cell.EffectiveRuns;
        if (runs.Count == 0 || runs.All(r => string.IsNullOrEmpty(r.Text)))
        {
            return;
        }

        float fontSize = Math.Max(1f, (float)(cell.FontSize * scale));
        float padding = 3f * scale;
        float available = rect.Width - (2 * padding);
        if (available <= 0)
        {
            return;
        }

        // Per-run fonts (family is whole-cell; weight/slant vary per run). Measure the total run
        // width so the whole string can be aligned within the cell.
        List<(CellRun Run, SKFont Font, float Width)> laid = new(runs.Count);
        float totalWidth = 0;
        foreach (CellRun run in runs)
        {
            if (string.IsNullOrEmpty(run.Text))
            {
                continue;
            }

            SKFont font = new(GetTypeface(typefaces, cell.FontFamily, run.Bold, run.Italic), fontSize);
            float width = font.MeasureText(run.Text);
            laid.Add((run, font, width));
            totalWidth += width;
        }

        if (laid.Count == 0)
        {
            return;
        }

        try
        {
            float x = cell.TextAlignment switch
            {
                TextAlignment.Center => rect.Left + ((rect.Width - totalWidth) / 2f),
                TextAlignment.Right => rect.Right - padding - totalWidth,
                _ => rect.Left + padding,
            };

            // Vertically center on the cell using the (uniform) cell font metrics.
            SKFontMetrics metrics = laid[0].Font.Metrics;
            float baseline = rect.MidY - ((metrics.Ascent + metrics.Descent) / 2f);

            canvas.Save();
            canvas.ClipRect(new SKRect(rect.Left + 0.5f, rect.Top, rect.Right - 0.5f, rect.Bottom));

            using SKPaint paint = new() { IsAntialias = true };
            foreach ((CellRun run, SKFont font, float width) in laid)
            {
                if (run.Background is { } runBg)
                {
                    using SKPaint runFill = new() { Color = ToSkColor(runBg), Style = SKPaintStyle.Fill };
                    canvas.DrawRect(new SKRect(x, rect.Top, x + width, rect.Bottom), runFill);
                }

                paint.Color = run.Foreground is { } fg ? ToSkColor(fg) : TextColor;
                canvas.DrawText(run.Text, x, baseline, font, paint);

                if (run.Underline)
                {
                    float underlineY = baseline + (metrics.UnderlinePosition ?? (fontSize * 0.12f));
                    canvas.DrawLine(x, underlineY, x + width, underlineY, paint);
                }

                if (run.Strikethrough)
                {
                    float strikeY = baseline + (metrics.StrikeoutPosition ?? -(fontSize * 0.3f));
                    canvas.DrawLine(x, strikeY, x + width, strikeY, paint);
                }

                x += width;
            }

            canvas.Restore();
        }
        finally
        {
            foreach ((CellRun _, SKFont font, float _) in laid)
            {
                font.Dispose();
            }
        }
    }

    private static SKTypeface GetTypeface(Dictionary<TypefaceKey, SKTypeface> cache, string? family, bool bold, bool italic)
    {
        TypefaceKey key = new(family ?? string.Empty, bold, italic);
        if (cache.TryGetValue(key, out SKTypeface? cached))
        {
            return cached;
        }

        SKTypeface typeface = SKTypeface.FromFamilyName(
            string.IsNullOrWhiteSpace(family) ? "Calibri" : family,
            bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
            SKFontStyleWidth.Normal,
            italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright)
            ?? SKTypeface.Default;

        cache[key] = typeface;
        return typeface;
    }

    private static SKColor ToSkColor(Color color) => new(color.R, color.G, color.B, color.A);

    private static double ColumnWidthPx(SpreadsheetViewModel sheet, int column) =>
        column < sheet.ColumnCount ? sheet.Columns[column].Width : sheet.DefaultColumnWidth;

    private static double RowHeightPx(SpreadsheetViewModel sheet, int row)
    {
        if (row >= sheet.RowCount)
        {
            return 22;
        }

        if (sheet.Rows[row].Height is { } explicitHeight)
        {
            return explicitHeight;
        }

        double maxFont = sheet.Rows[row].Cells.Count == 0
            ? CellViewModel.DefaultFontSize
            : sheet.Rows[row].Cells.Max(c => c.FontSize);
        return Math.Max(22, maxFont + 8);
    }

    private static int LastNonEmptyColumn(SpreadsheetViewModel sheet, int lastRow)
    {
        int last = 0;
        for (int r = 0; r <= lastRow && r < sheet.RowCount; r++)
        {
            RowViewModel row = sheet.Rows[r];
            for (int c = row.Cells.Count - 1; c > last; c--)
            {
                if (!string.IsNullOrEmpty(row.Cells[c].DisplayText))
                {
                    last = c;
                    break;
                }
            }
        }

        return last;
    }

    private static (Dictionary<(int, int), MergeRange> Anchors, HashSet<(int, int)> Covered) BuildMergeMaps(SpreadsheetViewModel sheet)
    {
        Dictionary<(int, int), MergeRange> anchors = new();
        HashSet<(int, int)> covered = new();
        foreach (MergeRange merge in sheet.Merges)
        {
            anchors[(merge.Row, merge.Column)] = merge;
            for (int r = merge.Row; r < merge.Row + merge.RowSpan; r++)
            {
                for (int c = merge.Column; c < merge.Column + merge.ColumnSpan; c++)
                {
                    if (r != merge.Row || c != merge.Column)
                    {
                        covered.Add((r, c));
                    }
                }
            }
        }

        return (anchors, covered);
    }

    private readonly record struct TypefaceKey(string Family, bool Bold, bool Italic);
}
