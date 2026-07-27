namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using ClosedXML.Excel;
using global::SpreadsheetEditor.Controls.Models;

// Reads charts embedded in an .xlsx package. ClosedXML drops charts, so this parses the OOXML parts
// directly: the worksheet drawing (anchor geometry + chart references) and each chart part (type +
// series value/category ranges + titles). Produces SheetCharts the grid renders live.
public static class XlsxCharts
{
    private const double EmuPerPixel = 9525.0; // 1px @ 96dpi

    private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rels = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // A chart's raw series as parsed from the chart XML: range formulas plus the cached
    // values/labels Excel stores in the chart (numCache/strCache).
    public sealed record RawSeries(string? TitleFormula, string? TitleLiteral, string? ValueFormula,
        string? CategoryFormula, IReadOnlyList<double> CachedValues, IReadOnlyList<string> CachedCategories);

    public static List<SheetChart> Read(byte[] bytes, IXLWorksheet? sheet, int sheetIndex = 0)
    {
        List<SheetChart> charts = new();
        if (sheet is null)
        {
            return charts;
        }

        try
        {
            using ZipArchive archive = new(new MemoryStream(bytes), ZipArchiveMode.Read);
            Dictionary<string, ZipArchiveEntry> entries = archive.Entries.ToDictionary(static e => e.FullName, StringComparer.OrdinalIgnoreCase);

            // This worksheet's own drawing only — otherwise charts from other sheets would appear on it.
            string? sheetDrawing = ResolveSheetDrawing(entries, sheetIndex);

            // No drawing wired to this sheet → it has no charts (don't fall back to all drawings, which
            // would duplicate other sheets' charts onto every tab).
            if (sheetDrawing is null)
            {
                return charts;
            }

            List<string> drawingNames = new() { sheetDrawing };

            foreach (string drawingName in drawingNames)
            {
                if (!entries.TryGetValue(drawingName, out ZipArchiveEntry? drawing))
                {
                    continue;
                }

                XDocument? drawingDoc = Load(drawing);
                if (drawingDoc is null)
                {
                    continue;
                }

                string relsName = $"xl/drawings/_rels/{Path.GetFileName(drawing.FullName)}.rels";
                Dictionary<string, string> relMap = entries.TryGetValue(relsName, out ZipArchiveEntry? relEntry) ? LoadRels(relEntry) : new();

                foreach (Anchor anchor in ParseAnchors(drawingDoc))
                {
                    if (!relMap.TryGetValue(anchor.ChartRelId, out string? target))
                    {
                        continue;
                    }

                    string chartPath = ResolvePath("xl/drawings", target);
                    if (!entries.TryGetValue(chartPath, out ZipArchiveEntry? chartEntry))
                    {
                        continue;
                    }

                    XDocument? chartDoc = Load(chartEntry);
                    if (chartDoc is null)
                    {
                        continue;
                    }

                    (ChartKind kind, List<RawSeries> rawSeries, bool showDataLabels, double? axisMinimum) = ParseChartPart(chartDoc);
                    SheetChart? built = BuildChart(kind, rawSeries, showDataLabels, axisMinimum, anchor, sheet);
                    if (built is not null)
                    {
                        charts.Add(built);
                    }
                }
            }
        }
        catch
        {
            // Never let a chart-parsing problem break opening the file.
        }

        return charts;
    }

    // ----- chart part -----

    // Parses a chart part into its kind, raw series, and whether data labels are shown.
    public static (ChartKind Kind, List<RawSeries> Series, bool ShowDataLabels, double? AxisMinimum) ParseChartPart(XDocument chartDoc)
    {
        XElement? plotArea = chartDoc.Descendants(C + "plotArea").FirstOrDefault();
        List<RawSeries> series = new();
        bool showDataLabels = chartDoc.Descendants(C + "showVal").Any(static e => e.Attribute("val")?.Value is "1" or "true");

        // Explicit value-axis minimum (<c:valAx><c:scaling><c:min>), if the chart pins one; otherwise null
        // → the builder defaults to 0 (Excel's automatic value-axis baseline).
        string? minText = chartDoc.Descendants(C + "valAx").FirstOrDefault()
            ?.Element(C + "scaling")?.Element(C + "min")?.Attribute("val")?.Value;
        double? axisMinimum = double.TryParse(minText, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedMin) ? parsedMin : null;

        if (plotArea is null)
        {
            return (ChartKind.Line, series, showDataLabels, axisMinimum);
        }

        // The first element whose name ends with "Chart" determines the type.
        XElement? typeElement = plotArea.Elements().FirstOrDefault(static e => e.Name.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        string? grouping = typeElement?.Element(C + "grouping")?.Attribute("val")?.Value;
        ChartKind kind = MapKind(typeElement?.Name.LocalName, grouping);

        foreach (XElement ser in (typeElement ?? plotArea).Elements(C + "ser"))
        {
            XElement? tx = ser.Element(C + "tx");
            string? titleFormula = tx?.Descendants(C + "f").FirstOrDefault()?.Value;
            string? titleLiteral = tx?.Descendants(C + "v").FirstOrDefault()?.Value;

            XElement? val = ser.Element(C + "val");
            XElement? cat = ser.Element(C + "cat");
            string? valueFormula = val?.Descendants(C + "f").FirstOrDefault()?.Value;
            string? categoryFormula = cat?.Descendants(C + "f").FirstOrDefault()?.Value;

            List<double> cachedValues = ReadCache(val).Select(static s =>
                double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double v) ? v : double.NaN).ToList();
            List<string> cachedCategories = ReadCache(cat);

            series.Add(new RawSeries(titleFormula, titleLiteral, valueFormula, categoryFormula, cachedValues, cachedCategories));
        }

        return (kind, series, showDataLabels, axisMinimum);
    }

    // Reads a c:val / c:cat element's cached point values (numCache or strCache) into a list
    // indexed by point index (gaps filled with empty strings).
    private static List<string> ReadCache(XElement? reference)
    {
        List<string> result = new();

        // Cached (num/strCache) or inline literal (num/strLit) point lists share the same shape.
        XElement? cache = reference?.Descendants(C + "numCache").FirstOrDefault()
                    ?? reference?.Descendants(C + "strCache").FirstOrDefault()
                    ?? reference?.Descendants(C + "numLit").FirstOrDefault()
                    ?? reference?.Descendants(C + "strLit").FirstOrDefault();
        if (cache is null)
        {
            return result;
        }

        int count = int.TryParse(cache.Element(C + "ptCount")?.Attribute("val")?.Value, out int n) ? n : 0;
        string[] values = new string[Math.Max(count, 0)];
        foreach (XElement pt in cache.Elements(C + "pt"))
        {
            if (!int.TryParse(pt.Attribute("idx")?.Value, out int idx) || idx < 0)
            {
                continue;
            }

            if (idx >= values.Length)
            {
                Array.Resize(ref values, idx + 1);
            }

            values[idx] = pt.Element(C + "v")?.Value ?? string.Empty;
        }

        result.AddRange(values.Select(static v => v ?? string.Empty));
        return result;
    }

    private static ChartKind MapKind(string? localName, string? grouping)
    {
        bool stacked = grouping is "stacked" or "percentStacked";
        return localName switch
        {
            "barChart" or "bar3DChart" => stacked ? ChartKind.StackedBar : ChartKind.Bar,
            "areaChart" or "area3DChart" => stacked ? ChartKind.StackedArea : ChartKind.Area,
            "scatterChart" => ChartKind.Scatter,
            "radarChart" => ChartKind.Radar,
            "doughnutChart" => ChartKind.Donut,
            "pieChart" or "pie3DChart" => ChartKind.Pie,
            _ => ChartKind.Line, // line / fallback
        };
    }

    // ----- range formulas -----

    // Parses an OOXML range formula like Sheet1!$D$18:$D$36 to a 0-based range.
    public static CellRange? ParseRange(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            return null;
        }

        // Take the first area (formulas may be multi-area, comma-separated), then drop its sheet
        // qualifier (everything up to the last '!', which also handles quoted sheet names).
        string firstArea = formula.Split(',')[0];
        int bang = firstArea.LastIndexOf('!');
        string refPart = (bang >= 0 ? firstArea[(bang + 1)..] : firstArea).Trim();
        if (refPart.Length == 0)
        {
            return null;
        }

        string[] ends = refPart.Split(':');
        (int Row, int Column)? a = ParseA1(ends[0]);
        (int Row, int Column)? b = ends.Length > 1 ? ParseA1(ends[1]) : a;
        if (a is null || b is null)
        {
            return null;
        }

        (int r1, int c1) = a.Value;
        (int r2, int c2) = b.Value;
        int minR = Math.Min(r1, r2);
        int minC = Math.Min(c1, c2);
        return new CellRange(minR, minC, Math.Abs(r2 - r1) + 1, Math.Abs(c2 - c1) + 1);
    }

    private static (int Row, int Column)? ParseA1(string cell)
    {
        string s = cell.Replace("$", "").Trim();
        int split = 0;
        while (split < s.Length && char.IsLetter(s[split]))
        {
            split++;
        }

        if (split == 0 || split == s.Length)
        {
            return null;
        }

        if (!int.TryParse(s[split..], NumberStyles.None, CultureInfo.InvariantCulture, out int row) || row < 1)
        {
            return null;
        }

        try
        {
            return (row - 1, CellAddress.LettersToColumn(s[..split]));
        }
        catch
        {
            return null;
        }
    }

    // ----- drawing anchors -----

    public sealed record Anchor(string ChartRelId, int FromCol, double FromColOffPx, int FromRow, double FromRowOffPx,
        bool HasTo, int ToCol, double ToColOffPx, int ToRow, double ToRowOffPx, double ExtWidthPx, double ExtHeightPx);

    private static List<Anchor> ParseAnchors(XDocument drawingDoc)
    {
        List<Anchor> anchors = new();
        foreach (XElement el in drawingDoc.Descendants().Where(e =>
                     e.Name == Xdr + "twoCellAnchor" || e.Name == Xdr + "oneCellAnchor"))
        {
            string? chartId = el.Descendants(C + "chart").FirstOrDefault()?.Attribute(R + "id")?.Value;
            if (chartId is null)
            {
                continue;
            }

            XElement? from = el.Element(Xdr + "from");
            if (from is null)
            {
                continue;
            }

            ReadMarker(from, out int fc, out double fco, out int fr, out double fro);

            if (el.Name == Xdr + "twoCellAnchor")
            {
                XElement? to = el.Element(Xdr + "to");
                if (to is not null)
                {
                    ReadMarker(to, out int tc, out double tco, out int tr, out double tro);
                    anchors.Add(new Anchor(chartId, fc, fco, fr, fro, true, tc, tco, tr, tro, 0, 0));
                    continue;
                }
            }

            // oneCellAnchor: from + explicit extent (EMU).
            XElement? ext = el.Element(Xdr + "ext");
            double cx = Px(ext?.Attribute("cx")?.Value);
            double cy = Px(ext?.Attribute("cy")?.Value);
            anchors.Add(new Anchor(chartId, fc, fco, fr, fro, false, 0, 0, 0, 0, cx, cy));
        }

        return anchors;
    }

    private static void ReadMarker(XElement marker, out int col, out double colOffPx, out int row, out double rowOffPx)
    {
        col = (int)(long.TryParse(marker.Element(Xdr + "col")?.Value, out long c) ? c : 0);
        row = (int)(long.TryParse(marker.Element(Xdr + "row")?.Value, out long r) ? r : 0);
        colOffPx = Px(marker.Element(Xdr + "colOff")?.Value);
        rowOffPx = Px(marker.Element(Xdr + "rowOff")?.Value);
    }

    private static double Px(string? emu) =>
        long.TryParse(emu, NumberStyles.Integer, CultureInfo.InvariantCulture, out long v) ? v / EmuPerPixel : 0;

    // ----- stitching -----

    private static SheetChart? BuildChart(ChartKind kind, List<RawSeries> rawSeries, bool showDataLabels, double? axisMinimum, Anchor anchor, IXLWorksheet sheet)
    {
        List<ChartSeriesSpec> series = new();
        CellRange? categories = null;
        IReadOnlyList<string>? cachedCategories = null;

        for (int i = 0; i < rawSeries.Count; i++)
        {
            RawSeries raw = rawSeries[i];
            CellRange? range = ParseRange(raw.ValueFormula);
            bool hasCache = raw.CachedValues.Count > 0;
            if (range is null && !hasCache)
            {
                continue;
            }

            string title = ResolveTitle(raw, sheet) ?? $"Series {i + 1}";
            series.Add(new ChartSeriesSpec(title, range ?? new CellRange(0, 0, 0, 0),
                hasCache ? raw.CachedValues : null));

            categories ??= ParseRange(raw.CategoryFormula);
            if (cachedCategories is null && raw.CachedCategories.Any(static s => !string.IsNullOrEmpty(s)))
            {
                cachedCategories = raw.CachedCategories;
            }
        }

        if (series.Count == 0)
        {
            return null;
        }

        // The grid sizes a two-cell-anchored chart to span from→to in its own layout (matches Excel's
        // cell coverage). A one-cell anchor keeps the explicit extent.
        return new SheetChart
        {
            Kind = kind,
            Series = series,
            Categories = categories,
            CachedCategories = cachedCategories,
            ShowDataLabels = showDataLabels,
            ShowLegend = series.Count > 1,
            // Honor an explicit axis minimum from the file; otherwise start at 0 like Excel's default.
            AxisMinimum = axisMinimum ?? 0,
            AnchorRow = anchor.FromRow,
            AnchorColumn = anchor.FromCol,
            OffsetX = anchor.FromColOffPx,
            OffsetY = anchor.FromRowOffPx,
            ToRow = anchor.HasTo ? anchor.ToRow : null,
            ToColumn = anchor.HasTo ? anchor.ToCol : null,
            ToOffsetX = anchor.ToColOffPx,
            ToOffsetY = anchor.ToRowOffPx,
            Width = anchor.ExtWidthPx > 0 ? anchor.ExtWidthPx : 380,
            Height = anchor.ExtHeightPx > 0 ? anchor.ExtHeightPx : 260,
        };
    }

    private static string? ResolveTitle(RawSeries raw, IXLWorksheet sheet)
    {
        if (!string.IsNullOrEmpty(raw.TitleLiteral))
        {
            return raw.TitleLiteral;
        }

        CellRange? range = ParseRange(raw.TitleFormula);
        if (range is null)
        {
            return null;
        }

        try
        {
            IXLCell cell = sheet.Cell(range.Value.Row + 1, range.Value.Column + 1);
            string text = cell.GetFormattedString();
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }
        catch
        {
            return null;
        }
    }

    // Resolves the drawing part belonging to the first worksheet (workbook → first sheet →
    // its drawing rel), or null if it can't be determined.
    private static string? ResolveSheetDrawing(Dictionary<string, ZipArchiveEntry> entries, int sheetIndex)
    {
        try
        {
            if (!entries.TryGetValue("xl/workbook.xml", out ZipArchiveEntry? wbEntry) ||
                !entries.TryGetValue("xl/_rels/workbook.xml.rels", out ZipArchiveEntry? wbRelsEntry))
            {
                return null;
            }

            XDocument? wb = Load(wbEntry);

            // The Nth <sheet> in workbook.xml is the Nth worksheet (tab order matches ClosedXML's).
            List<XElement>? sheetEls = wb?.Descendants(Main + "sheet").ToList();
            if (sheetEls is null || sheetIndex < 0 || sheetIndex >= sheetEls.Count)
            {
                return null;
            }

            string? sheetRid = sheetEls[sheetIndex].Attribute(R + "id")?.Value;
            if (sheetRid is null)
            {
                return null;
            }

            Dictionary<string, string> wbRels = LoadRels(wbRelsEntry);
            if (!wbRels.TryGetValue(sheetRid, out string? sheetTarget))
            {
                return null;
            }

            string sheetPath = ResolvePath("xl", sheetTarget); // e.g. xl/worksheets/sheet1.xml

            if (!entries.TryGetValue(sheetPath, out ZipArchiveEntry? sheetEntry))
            {
                return null;
            }

            XDocument? sheetDoc = Load(sheetEntry);
            string? drawingRid = sheetDoc?.Descendants(Main + "drawing").FirstOrDefault()?.Attribute(R + "id")?.Value;
            if (drawingRid is null)
            {
                return null;
            }

            string sheetDir = sheetPath[..sheetPath.LastIndexOf('/')];
            string sheetRelsName = $"{sheetDir}/_rels/{Path.GetFileName(sheetPath)}.rels";
            if (!entries.TryGetValue(sheetRelsName, out ZipArchiveEntry? sheetRelsEntry))
            {
                return null;
            }

            Dictionary<string, string> sheetRels = LoadRels(sheetRelsEntry);
            return sheetRels.TryGetValue(drawingRid, out string? drawingTarget)
                ? ResolvePath(sheetDir, drawingTarget)
                : null;
        }
        catch
        {
            return null;
        }
    }

    // ----- zip helpers -----

    private static XDocument? Load(ZipArchiveEntry entry)
    {
        try
        {
            using Stream s = entry.Open();
            return XDocument.Load(s);
        }
        catch
        {
            return null;
        }
    }

    private static Dictionary<string, string> LoadRels(ZipArchiveEntry entry)
    {
        Dictionary<string, string> map = new(StringComparer.OrdinalIgnoreCase);
        XDocument? doc = Load(entry);
        if (doc is null)
        {
            return map;
        }

        foreach (XElement rel in doc.Descendants(Rels + "Relationship"))
        {
            string? id = rel.Attribute("Id")?.Value;
            string? target = rel.Attribute("Target")?.Value;
            if (id is not null && target is not null)
            {
                map[id] = target;
            }
        }

        return map;
    }

    // Resolves a relationship target (possibly "../charts/chart1.xml") against a base dir.
    private static string ResolvePath(string baseDir, string target)
    {
        if (target.StartsWith('/'))
        {
            return target.TrimStart('/');
        }

        List<string> parts = new(baseDir.Split('/', StringSplitOptions.RemoveEmptyEntries));
        foreach (string seg in target.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (seg == ".")
            {
                continue;
            }

            if (seg == "..")
            {
                if (parts.Count > 0)
                {
                    parts.RemoveAt(parts.Count - 1);
                }
            }
            else
            {
                parts.Add(seg);
            }
        }

        return string.Join('/', parts);
    }
}
