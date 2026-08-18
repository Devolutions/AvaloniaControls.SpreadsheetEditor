namespace SpreadsheetEditor.Controls.Printing;

using System;
using System.Collections.Generic;
using System.Globalization;
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
// and rows flow onto as many pages as needed. Charts are composited from pre-rendered snapshots
// (see ChartImage) positioned in grid coordinates; images are not rendered.
public static class SpreadsheetPdfExporter
{
    // US Letter, portrait, at 72 pt/inch (PDF's native unit).
    private const float PageWidth = 612f;

    private const float PageHeight = 792f;

    private const float Margin = 36f;

    // Space reserved at the top of each page above the grid. The sheet-name title is not printed, so
    // this is just a small breathing gap below the page margin.
    private const float TitleHeight = 0f;

    // Screen sizes (column widths, row heights, font sizes) are in device px at 96 dpi;
    // PDF is 72 pt/inch, so 1 px = 0.75 pt.
    private const double PxToPt = 72.0 / 96.0;

    private static readonly SKColor GridLineColor = new(0xC8, 0xC8, 0xC8);

    private static readonly SKColor TextColor = SKColors.Black;

    public static void Export(SpreadsheetViewModel sheet, Stream stream, IReadOnlyList<ChartImage>? charts = null)
    {
        if (sheet is null)
        {
            throw new ArgumentNullException(nameof(sheet));
        }

        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        PageLayout layout = ComputeLayout(sheet);
        Dictionary<TypefaceKey, SKTypeface> typefaces = new();
        try
        {
            SKDocumentPdfMetadata metadata = new() { Title = sheet.Name };
            using SKDocument document = SKDocument.CreatePdf(stream, metadata);
            for (int p = 0; p < layout.Pages.Count; p++)
            {
                SKCanvas canvas = document.BeginPage(PageWidth, PageHeight);
                try
                {
                    DrawPage(canvas, sheet, charts, layout, p, typefaces);
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

    // Renders each page to a raster bitmap (BGRA, top-down) at the given DPI. Used by the Windows print
    // path, which blits these onto the printer's device context. Caller owns/disposes the bitmaps.
    internal static List<SKBitmap> RenderPages(SpreadsheetViewModel sheet, IReadOnlyList<ChartImage>? charts, float dpi)
    {
        PageLayout layout = ComputeLayout(sheet);
        float k = dpi / 72f;
        int widthPx = Math.Max(1, (int)Math.Ceiling(PageWidth * k));
        int heightPx = Math.Max(1, (int)Math.Ceiling(PageHeight * k));

        List<SKBitmap> bitmaps = new();
        Dictionary<TypefaceKey, SKTypeface> typefaces = new();
        try
        {
            for (int p = 0; p < layout.Pages.Count; p++)
            {
                SKBitmap bitmap = new(new SKImageInfo(widthPx, heightPx, SKColorType.Bgra8888, SKAlphaType.Premul));
                using (SKCanvas canvas = new(bitmap))
                {
                    canvas.Clear(SKColors.White);
                    canvas.Scale(k);
                    DrawPage(canvas, sheet, charts, layout, p, typefaces);
                }

                bitmaps.Add(bitmap);
            }
        }
        finally
        {
            foreach (SKTypeface typeface in typefaces.Values)
            {
                typeface.Dispose();
            }
        }

        return bitmaps;
    }

    // Computes the shared layout (fit-to-width scale, column/row geometry, page row-ranges) once, so
    // both the PDF and raster-print paths render identically.
    private static PageLayout ComputeLayout(SpreadsheetViewModel sheet)
    {
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

        float[] columnLeft = new float[lastColumn + 2];
        columnLeft[0] = Margin;
        for (int c = 0; c <= lastColumn; c++)
        {
            columnLeft[c + 1] = columnLeft[c] + (float)(ColumnWidthPx(sheet, c) * scale);
        }

        float[] rowHeightPt = new float[lastRow + 1];
        double[] rowTopPx = new double[lastRow + 2];
        for (int r = 0; r <= lastRow; r++)
        {
            double heightPx = RowHeightPx(sheet, r);
            rowHeightPt[r] = (float)(heightPx * scale);
            rowTopPx[r + 1] = rowTopPx[r] + heightPx;
        }

        float gridTop = Margin + TitleHeight;
        float contentBottom = PageHeight - Margin;

        List<(int Start, int End)> pages = new();
        int cursor = 0;
        while (cursor <= lastRow)
        {
            float y = gridTop;
            int pageStart = cursor;
            while (cursor <= lastRow && (cursor == pageStart || y + rowHeightPt[cursor] <= contentBottom))
            {
                y += rowHeightPt[cursor];
                cursor++;
            }

            pages.Add((pageStart, cursor));
        }

        (Dictionary<(int, int), (int, int)> spills, HashSet<(int, int)> noLeft, HashSet<(int, int)> noRight) =
            ComputeSpills(sheet, lastRow, lastColumn, covered);

        return new PageLayout
        {
            LastRow = lastRow,
            LastColumn = lastColumn,
            Anchors = anchors,
            Covered = covered,
            Spills = spills,
            NoLeftBorder = noLeft,
            NoRightBorder = noRight,
            Scale = scale,
            ColumnLeft = columnLeft,
            RowHeightPt = rowHeightPt,
            RowTopPx = rowTopPx,
            GridTop = gridTop,
            ContentBottom = contentBottom,
            Pages = pages,
        };
    }

    // Draws a single page (title, cells, data bars, charts) onto the given canvas in page points.
    private static void DrawPage(SKCanvas canvas, SpreadsheetViewModel sheet, IReadOnlyList<ChartImage>? charts, PageLayout layout, int pageIndex, Dictionary<TypefaceKey, SKTypeface> typefaces)
    {
        (int pageStart, int pageEnd) = layout.Pages[pageIndex];
        bool lastPage = pageIndex == layout.Pages.Count - 1;
        double pageTopPx = layout.RowTopPx[pageStart];
        double pageBottomPx = layout.RowTopPx[pageEnd];

        // Clip so a merged block or chart spanning past the page bottom is trimmed cleanly.
        canvas.Save();
        canvas.ClipRect(new SKRect(Margin, layout.GridTop, PageWidth - Margin, layout.ContentBottom));

        float rowTop = layout.GridTop;
        for (int r = pageStart; r < pageEnd; r++)
        {
            for (int c = 0; c <= layout.LastColumn; c++)
            {
                if (layout.Covered.Contains((r, c)))
                {
                    continue;
                }

                CellViewModel? cell = sheet.GetCell(new CellAddress(r, c));

                int columnSpan = 1;
                int rowSpan = 1;
                if (layout.Anchors.TryGetValue((r, c), out MergeRange span))
                {
                    columnSpan = span.ColumnSpan;
                    rowSpan = span.RowSpan;
                }

                float cellLeft = layout.ColumnLeft[c];
                float cellRight = layout.ColumnLeft[Math.Min(c + columnSpan, layout.LastColumn + 1)];
                float cellBottom = rowTop;
                for (int rr = r; rr < r + rowSpan && rr <= layout.LastRow; rr++)
                {
                    cellBottom += layout.RowHeightPt[rr];
                }

                SKRect rect = new(cellLeft, rowTop, cellRight, cellBottom);

                // When the cell spills, its text is drawn across the empty neighbours (a wider rect)
                // and the shared gridlines are dropped so the run reads as continuous — matching screen.
                SKRect textRect = rect;
                if (layout.Spills.TryGetValue((r, c), out (int Left, int Right) spill))
                {
                    float spanLeft = layout.ColumnLeft[c - spill.Left];
                    float spanRight = layout.ColumnLeft[Math.Min(c + spill.Right + 1, layout.LastColumn + 1)];
                    textRect = new SKRect(spanLeft, rect.Top, spanRight, rect.Bottom);
                }

                DrawCell(canvas, cell, rect, textRect, (float)layout.Scale, typefaces, FindDataBar(sheet, r, c),
                    layout.NoLeftBorder.Contains((r, c)), layout.NoRightBorder.Contains((r, c)));
            }

            rowTop += layout.RowHeightPt[r];
        }

        if (charts is not null)
        {
            foreach (ChartImage chart in charts)
            {
                if (chart.Png is null || chart.Png.Length == 0 || chart.Width <= 0 || chart.Height <= 0)
                {
                    continue;
                }

                if (chart.Top < pageTopPx || (!lastPage && chart.Top >= pageBottomPx))
                {
                    continue;
                }

                DrawChart(canvas, chart, layout.GridTop, pageTopPx, layout.Scale);
            }
        }

        canvas.Restore();
    }

    private static void DrawChart(SKCanvas canvas, ChartImage chart, float gridTop, double pageTopPx, double scale)
    {
        using SKImage? image = SKImage.FromEncodedData(chart.Png);
        if (image is null)
        {
            return;
        }

        float left = Margin + (float)(chart.Left * scale);
        float top = gridTop + (float)((chart.Top - pageTopPx) * scale);
        float right = left + (float)(chart.Width * scale);
        float bottom = top + (float)(chart.Height * scale);
        SKRect dest = new(left, top, right, bottom);
        canvas.DrawImage(image, dest, new SKSamplingOptions(SKFilterMode.Linear, SKMipmapMode.None));
    }

    private static void DrawCell(SKCanvas canvas, CellViewModel? cell, SKRect rect, SKRect textRect, float scale, Dictionary<TypefaceKey, SKTypeface> typefaces, SheetDataBar? dataBar, bool omitLeftBorder, bool omitRightBorder)
    {
        // Background fill.
        if (cell?.BackgroundColor is { } bg)
        {
            using SKPaint fill = new() { Color = ToSkColor(bg), Style = SKPaintStyle.Fill };
            canvas.DrawRect(rect, fill);
        }

        // Conditional-formatting data bar: a horizontal fill proportional to the cell value between
        // Min and Max (drawn behind the text, as on screen).
        if (dataBar is { } bar)
        {
            double value = ParseNumber(cell?.DisplayText) ?? bar.Min;
            double fraction = bar.Max > bar.Min ? Math.Clamp((value - bar.Min) / (bar.Max - bar.Min), 0, 1) : 0;
            float margin = scale;
            float barRight = rect.Left + margin + (float)((rect.Width - (2 * margin)) * fraction);
            if (barRight > rect.Left + margin)
            {
                using SKPaint barPaint = new() { Color = ToSkColor(bar.Color), Style = SKPaintStyle.Fill, IsAntialias = true };
                canvas.DrawRect(new SKRect(rect.Left + margin, rect.Top + margin, barRight, rect.Bottom - margin), barPaint);
            }
        }

        // Cell border (thin gray, matching the HTML export's grid lines). A vertical edge shared with a
        // cell that this one spills into (or that spills into this one) is dropped so the gridline
        // between them disappears, exactly as on screen.
        using (SKPaint border = new() { Color = GridLineColor, Style = SKPaintStyle.Stroke, StrokeWidth = 0.5f, IsAntialias = true })
        {
            canvas.DrawLine(rect.Left, rect.Top, rect.Right, rect.Top, border);       // top
            canvas.DrawLine(rect.Left, rect.Bottom, rect.Right, rect.Bottom, border); // bottom
            if (!omitLeftBorder)
            {
                canvas.DrawLine(rect.Left, rect.Top, rect.Left, rect.Bottom, border);
            }
            if (!omitRightBorder)
            {
                canvas.DrawLine(rect.Right, rect.Top, rect.Right, rect.Bottom, border);
            }
        }

        if (cell is null || dataBar is { ShowValue: false })
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
        float available = textRect.Width - (2 * padding);
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
                TextAlignment.Center => textRect.Left + ((textRect.Width - totalWidth) / 2f),
                TextAlignment.Right => textRect.Right - padding - totalWidth,
                _ => textRect.Left + padding,
            };

            // Vertically center on the cell using the (uniform) cell font metrics.
            SKFontMetrics metrics = laid[0].Font.Metrics;
            float baseline = textRect.MidY - ((metrics.Ascent + metrics.Descent) / 2f);

            canvas.Save();
            canvas.ClipRect(new SKRect(textRect.Left + 0.5f, textRect.Top, textRect.Right - 0.5f, textRect.Bottom));

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

    private static SheetDataBar? FindDataBar(SpreadsheetViewModel sheet, int row, int column)
    {
        foreach (SheetDataBar bar in sheet.DataBars)
        {
            if (row >= bar.Range.Row && row < bar.Range.Row + bar.Range.RowSpan &&
                column >= bar.Range.Column && column < bar.Range.Column + bar.Range.ColumnSpan)
            {
                return bar;
            }
        }

        return null;
    }

    // Parses a displayed value into a number (mirrors the grid): tolerates $, %, thousands separators,
    // and parenthesised negatives.
    private static double? ParseNumber(string? s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return null;
        }

        string t = s.Trim();
        bool negative = t.StartsWith('(') && t.EndsWith(')');
        if (negative)
        {
            t = t[1..^1];
        }

        bool percent = t.EndsWith('%');
        if (percent)
        {
            t = t[..^1];
        }

        t = t.Replace("$", string.Empty).Replace(",", string.Empty).Replace("€", string.Empty).Replace("£", string.Empty).Replace(" ", string.Empty).Trim();
        if (!double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
        {
            return null;
        }

        if (percent)
        {
            v /= 100;
        }

        return negative ? -v : v;
    }

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

    // Mirrors SpreadsheetGrid's on-screen text spill: a text cell (not number-formatted, not an error,
    // no data bar) overflows into adjacent empty cells, and the gridlines between it and those cells
    // are removed. NoRightBorder/NoLeftBorder mark the internal cell edges to drop so the spanned run
    // reads as continuous — an edge is only removed when BOTH neighbours agree it is internal.
    private static (Dictionary<(int, int), (int Left, int Right)> Spills, HashSet<(int, int)> NoLeft, HashSet<(int, int)> NoRight)
        ComputeSpills(SpreadsheetViewModel sheet, int lastRow, int lastColumn, HashSet<(int, int)> covered)
    {
        Dictionary<(int, int), (int, int)> spills = new();
        HashSet<(int, int)> noLeft = new();
        HashSet<(int, int)> noRight = new();
        for (int r = 0; r <= lastRow && r < sheet.RowCount; r++)
        {
            RowViewModel row = sheet.Rows[r];
            for (int c = 0; c <= lastColumn && c < row.Cells.Count; c++)
            {
                if (covered.Contains((r, c)))
                {
                    continue;
                }

                (int left, int right) = ComputeSpill(sheet, r, c, row.Cells[c], lastColumn, covered);
                if (left + right == 0)
                {
                    continue;
                }

                spills[(r, c)] = (left, right);
                for (int k = c - left; k < c + right; k++)
                {
                    noRight.Add((r, k));
                }
                for (int k = c - left + 1; k <= c + right; k++)
                {
                    noLeft.Add((r, k));
                }
            }
        }

        return (spills, noLeft, noRight);
    }

    private static (int Left, int Right) ComputeSpill(SpreadsheetViewModel sheet, int row, int column, CellViewModel cell, int lastColumn, HashSet<(int, int)> covered)
    {
        if (cell.NumberFormat is not null)
        {
            return (0, 0); // number-formatted (e.g. currency) cells don't spill — Excel shows ####
        }

        string display = cell.DisplayText;
        if (string.IsNullOrWhiteSpace(display) ||
            double.TryParse(display, NumberStyles.Any, CultureInfo.InvariantCulture, out _) ||
            IsErrorValue(display) ||
            FindDataBar(sheet, row, column) is not null)
        {
            return (0, 0);
        }

        TextAlignment align = cell.TextAlignment;
        int rightRun = align != TextAlignment.Right ? EmptyRun(sheet, row, column, +1, lastColumn, covered) : 0;
        int leftRun = align != TextAlignment.Left ? EmptyRun(sheet, row, column, -1, lastColumn, covered) : 0;
        if (rightRun == 0 && leftRun == 0)
        {
            return (0, 0);
        }

        // Overflow only as far as the text is wide — drop just the gridlines under the overflowing
        // text, not every empty cell in the row (mirrors SpreadsheetGrid).
        double overflow = MeasureTextWidth(cell) - (ColumnWidthPx(sheet, column) - CellTextPadding);
        if (overflow <= 0)
        {
            return (0, 0);
        }

        return align switch
        {
            TextAlignment.Right => (ColumnsToCover(sheet, column, -1, overflow, leftRun, lastColumn), 0),
            TextAlignment.Center => (ColumnsToCover(sheet, column, -1, overflow / 2, leftRun, lastColumn), ColumnsToCover(sheet, column, +1, overflow / 2, rightRun, lastColumn)),
            _ => (0, ColumnsToCover(sheet, column, +1, overflow, rightRun, lastColumn)),
        };
    }

    // Total horizontal text padding (both sides), in px, matching DrawCell's per-side padding.
    private const double CellTextPadding = 6;

    // Measures a cell's text width in px (per-run typefaces), for deciding how far it overflows.
    private static double MeasureTextWidth(CellViewModel cell)
    {
        double width = 0;
        foreach (CellRun run in cell.EffectiveRuns)
        {
            if (string.IsNullOrEmpty(run.Text))
            {
                continue;
            }
            SKTypeface typeface = SKTypeface.FromFamilyName(
                string.IsNullOrWhiteSpace(cell.FontFamily) ? "Calibri" : cell.FontFamily,
                run.Bold ? SKFontStyleWeight.Bold : SKFontStyleWeight.Normal,
                SKFontStyleWidth.Normal,
                run.Italic ? SKFontStyleSlant.Italic : SKFontStyleSlant.Upright) ?? SKTypeface.Default;
            using (SKFont font = new(typeface, (float)cell.FontSize))
            {
                width += font.MeasureText(run.Text);
            }

            if (!ReferenceEquals(typeface, SKTypeface.Default))
            {
                typeface.Dispose();
            }
        }

        return width;
    }

    private static int ColumnsToCover(SpreadsheetViewModel sheet, int column, int direction, double overflow, int maxRun, int lastColumn)
    {
        int covered = 0;
        double accumulated = 0;
        for (int i = 1; i <= maxRun; i++)
        {
            int col = column + (direction * i);
            if (col < 0 || col > lastColumn)
            {
                break;
            }
            covered = i;
            accumulated += ColumnWidthPx(sheet, col);
            if (accumulated >= overflow)
            {
                break;
            }
        }

        return covered;
    }

    private static int EmptyRun(SpreadsheetViewModel sheet, int row, int column, int direction, int lastColumn, HashSet<(int, int)> covered)
    {
        int count = 0;
        RowViewModel r = sheet.Rows[row];
        for (int c = column + direction; c >= 0 && c <= lastColumn && c < r.Cells.Count; c += direction)
        {
            if (covered.Contains((row, c)) || FindDataBar(sheet, row, c) is not null)
            {
                break;
            }
            if (!string.IsNullOrWhiteSpace(r.Cells[c].DisplayText))
            {
                break;
            }
            count++;
        }

        return count;
    }

    private static bool IsErrorValue(string display) =>
        display.Length > 1 && display[0] == '#' && (display[^1] is '!' or '?' || display == "#N/A");

    private readonly record struct TypefaceKey(string Family, bool Bold, bool Italic);

    // Precomputed page geometry shared by the PDF and raster-print renderers.
    private sealed class PageLayout
    {
        public int LastRow { get; init; }

        public int LastColumn { get; init; }

        public Dictionary<(int Row, int Column), MergeRange> Anchors { get; init; } = new();

        public HashSet<(int Row, int Column)> Covered { get; init; } = new();

        // Cells whose text overflows into empty neighbours, and which cell edges to drop so the
        // gridline between a spilling cell and the cell it spills into disappears (as on screen).
        public Dictionary<(int Row, int Column), (int Left, int Right)> Spills { get; init; } = new();

        public HashSet<(int Row, int Column)> NoLeftBorder { get; init; } = new();

        public HashSet<(int Row, int Column)> NoRightBorder { get; init; } = new();

        public double Scale { get; init; }

        public float[] ColumnLeft { get; init; } = Array.Empty<float>();

        public float[] RowHeightPt { get; init; } = Array.Empty<float>();

        public double[] RowTopPx { get; init; } = Array.Empty<double>();

        public float GridTop { get; init; }

        public float ContentBottom { get; init; }

        public List<(int Start, int End)> Pages { get; init; } = new();
    }
}

// A pre-rendered snapshot of one chart plus its position/size in grid px (the same coordinate
// system as the on-screen grid: Left/Top are absolute offsets from the first cell). Produced on the
// UI thread by SpreadsheetGrid.RenderChartImages and composited into the PDF by the exporter.
public readonly record struct ChartImage(byte[] Png, double Left, double Top, double Width, double Height);
