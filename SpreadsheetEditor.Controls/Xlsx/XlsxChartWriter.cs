namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;

using global::SpreadsheetEditor.Controls.Models;

// Writes charts into a ClosedXML-produced .xlsx by hand-authoring the OOXML chart/drawing parts and
// wiring the relationships — ClosedXML itself cannot create charts. Pure managed (System.IO.Compression
// + System.Xml.Linq, the same stack the chart reader uses), so the library stays a portable Avalonia
// component. The generated parts mirror what XlsxCharts.Read expects, so charts survive save/reopen.
internal static class XlsxChartWriter
{
    private const double EmuPerPixel = 9525.0;

    private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Rels = "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace Ct = "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    private const string DrawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
    private const string ChartContentType = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    private const string DrawingRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
    private const string ChartRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";

    // Injects the charts into the given .xlsx bytes and returns the new bytes. On any failure the
    // original bytes are returned unchanged (charts are best-effort, never break a save).
    public static byte[] Inject(byte[] xlsxBytes, IReadOnlyList<ChartWriteData> charts, int sheetIndex = 0)
    {
        if (charts is null || charts.Count == 0)
        {
            return xlsxBytes;
        }

        try
        {
            // Read the whole package into memory so we can freely add/replace parts.
            Dictionary<string, byte[]> parts = ReadParts(xlsxBytes);

            string? worksheetPart = parts.Keys
                .Where(k => k.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                    k.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                .OrderBy(WorksheetPartNumber)
                .ElementAtOrDefault(sheetIndex);
            if (worksheetPart is null)
            {
                return xlsxBytes;
            }

            string sheetFile = Path.GetFileName(worksheetPart);
            string sheetRelsPart = $"xl/worksheets/_rels/{sheetFile}.rels";

            // Build one drawing for this worksheet. Sheet-qualified part names let repeated calls add
            // charts to multiple worksheets without replacing an earlier sheet's drawing/chart parts.
            int partNumber = sheetIndex + 1;
            string drawingPart = $"xl/drawings/drawing{partNumber}.xml";
            string drawingRelsPart = $"xl/drawings/_rels/drawing{partNumber}.xml.rels";
            const string drawingRelId = "rIdSseChartDrawing";
            XElement wsDr = new(Xdr + "wsDr");
            List<(string Id, string Target)> drawingRels = new();
            for (int i = 0; i < charts.Count; i++)
            {
                string chartRelId = $"rId{i + 1}";
                string chartFile = $"chart{partNumber}_{i + 1}.xml";
                string chartPart = $"xl/charts/{chartFile}";
                parts[chartPart] = ToBytes(BuildChartXml(charts[i]));
                drawingRels.Add((chartRelId, $"../charts/{chartFile}"));
                wsDr.Add(BuildAnchor(charts[i], chartRelId));
            }

            parts[drawingPart] = ToBytes(new XDocument(wsDr));
            parts[drawingRelsPart] = ToBytes(BuildRels(drawingRels.Select(r => (r.Id, ChartRelType, r.Target))));

            // Point the worksheet at the drawing (element + relationship).
            parts[worksheetPart] = ToBytes(AddDrawingElement(parts[worksheetPart], drawingRelId));
            parts[sheetRelsPart] = ToBytes(AddRelationship(parts.TryGetValue(sheetRelsPart, out byte[]? existing) ? existing : null, drawingRelId, DrawingRelType, $"../drawings/drawing{partNumber}.xml"));

            // Register content types for the new parts.
            List<(string PartName, string ContentType)> overrides = new() { ($"/{drawingPart}", DrawingContentType) };
            for (int i = 0; i < charts.Count; i++)
            {
                overrides.Add(($"/xl/charts/chart{partNumber}_{i + 1}.xml", ChartContentType));
            }

            parts["[Content_Types].xml"] = ToBytes(AddContentTypes(parts["[Content_Types].xml"], overrides));

            return WriteParts(parts);
        }
        catch
        {
            return xlsxBytes;
        }
    }

    private static int WorksheetPartNumber(string path)
    {
        string name = Path.GetFileNameWithoutExtension(path);
        return name.StartsWith("sheet", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(name[5..], out int number)
            ? number
            : int.MaxValue;
    }

    private static Dictionary<string, byte[]> ReadParts(byte[] bytes)
    {
        Dictionary<string, byte[]> parts = new(StringComparer.OrdinalIgnoreCase);
        using ZipArchive archive = new(new MemoryStream(bytes), ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries)
        {
            using Stream s = entry.Open();
            using MemoryStream buffer = new();
            s.CopyTo(buffer);
            parts[entry.FullName] = buffer.ToArray();
        }

        return parts;
    }

    private static byte[] WriteParts(Dictionary<string, byte[]> parts)
    {
        using MemoryStream output = new();
        using (ZipArchive archive = new(output, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (KeyValuePair<string, byte[]> part in parts)
            {
                ZipArchiveEntry entry = archive.CreateEntry(part.Key, CompressionLevel.Optimal);
                using Stream s = entry.Open();
                s.Write(part.Value, 0, part.Value.Length);
            }
        }

        return output.ToArray();
    }

    private static XElement BuildAnchor(ChartWriteData chart, string chartRelId)
    {
        long cx = (long)Math.Round(Math.Max(1, chart.WidthPx) * EmuPerPixel);
        long cy = (long)Math.Round(Math.Max(1, chart.HeightPx) * EmuPerPixel);
        long colOff = (long)Math.Round(Math.Max(0, chart.OffsetXPx) * EmuPerPixel);
        long rowOff = (long)Math.Round(Math.Max(0, chart.OffsetYPx) * EmuPerPixel);

        return new XElement(
            Xdr + "oneCellAnchor",
            new XElement(
                Xdr + "from",
                new XElement(Xdr + "col", chart.AnchorColumn),
                new XElement(Xdr + "colOff", colOff),
                new XElement(Xdr + "row", chart.AnchorRow),
                new XElement(Xdr + "rowOff", rowOff)),
            new XElement(Xdr + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy)),
            new XElement(
                Xdr + "graphicFrame",
                new XAttribute("macro", string.Empty),
                new XElement(
                    Xdr + "nvGraphicFramePr",
                    new XElement(Xdr + "cNvPr", new XAttribute("id", 2), new XAttribute("name", "Chart")),
                    new XElement(Xdr + "cNvGraphicFramePr")),
                new XElement(
                    Xdr + "xfrm",
                    new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(A + "ext", new XAttribute("cx", cx), new XAttribute("cy", cy))),
                new XElement(
                    A + "graphic",
                    new XElement(
                        A + "graphicData",
                        new XAttribute("uri", C.NamespaceName),
                        new XElement(C + "chart", new XAttribute(R + "id", chartRelId))))),
            new XElement(Xdr + "clientData"));
    }

    private static XDocument BuildChartXml(ChartWriteData chart)
    {
        (string element, string? barDir, string? grouping, bool hasAxes) = ChartElement(chart.Kind);

        XElement plotType = new(C + element);
        if (barDir is not null)
        {
            plotType.Add(new XElement(C + "barDir", new XAttribute("val", barDir)));
        }

        if (grouping is not null)
        {
            plotType.Add(new XElement(C + "grouping", new XAttribute("val", grouping)));
        }

        plotType.Add(new XElement(C + "varyColors", new XAttribute("val", chart.Kind is ChartKind.Pie or ChartKind.Donut ? 1 : 0)));

        for (int i = 0; i < chart.Series.Count; i++)
        {
            plotType.Add(BuildSeries(chart, chart.Series[i], i));
        }

        if (hasAxes)
        {
            plotType.Add(new XElement(C + "axId", new XAttribute("val", 1)));
            plotType.Add(new XElement(C + "axId", new XAttribute("val", 2)));
        }

        XElement plotArea = new(C + "plotArea", new XElement(C + "layout"), plotType);
        if (hasAxes)
        {
            plotArea.Add(BuildCategoryAxis());
            plotArea.Add(BuildValueAxis(chart.AxisMinimum));
        }

        XElement chartEl = new(C + "chart");
        if (!string.IsNullOrWhiteSpace(chart.Title))
        {
            chartEl.Add(new XElement(
                C + "title",
                new XElement(
                    C + "tx",
                    new XElement(
                        C + "rich",
                        new XElement(A + "bodyPr"),
                        new XElement(A + "lstStyle"),
                        new XElement(A + "p", new XElement(A + "r", new XElement(A + "t", chart.Title)))))));
            chartEl.Add(new XElement(C + "autoTitleDeleted", new XAttribute("val", 0)));
        }
        else
        {
            chartEl.Add(new XElement(C + "autoTitleDeleted", new XAttribute("val", 1)));
        }

        chartEl.Add(plotArea);
        if (chart.ShowLegend)
        {
            chartEl.Add(new XElement(C + "legend", new XElement(C + "legendPos", new XAttribute("val", "b"))));
        }

        chartEl.Add(new XElement(C + "plotVisOnly", new XAttribute("val", 1)));

        XElement chartSpace = new(
            C + "chartSpace",
            new XAttribute(XNamespace.Xmlns + "c", C.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "a", A.NamespaceName),
            new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName),
            chartEl);

        return new XDocument(chartSpace);
    }

    private static XElement BuildSeries(ChartWriteData chart, ChartWriteSeries series, int index)
    {
        XElement ser = new(
            C + "ser",
            new XElement(C + "idx", new XAttribute("val", index)),
            new XElement(C + "order", new XAttribute("val", index)),
            new XElement(
                C + "tx",
                new XElement(
                    C + "strRef",
                    new XElement(C + "f", string.Empty),
                    StrCache(new[] { series.Title }))));

        if (chart.ShowDataLabels)
        {
            ser.Add(new XElement(
                C + "dLbls",
                new XElement(C + "showLegendKey", new XAttribute("val", 0)),
                new XElement(C + "showVal", new XAttribute("val", 1)),
                new XElement(C + "showCatName", new XAttribute("val", 0)),
                new XElement(C + "showSerName", new XAttribute("val", 0)),
                new XElement(C + "showPercent", new XAttribute("val", 0)),
                new XElement(C + "showBubbleSize", new XAttribute("val", 0))));
        }

        if (!string.IsNullOrEmpty(chart.CategoryFormula) || chart.CategoryLabels.Count > 0)
        {
            ser.Add(new XElement(
                C + "cat",
                new XElement(
                    C + "strRef",
                    new XElement(C + "f", chart.CategoryFormula ?? string.Empty),
                    StrCache(chart.CategoryLabels))));
        }

        ser.Add(new XElement(
            C + "val",
            new XElement(
                C + "numRef",
                new XElement(C + "f", series.ValueFormula),
                NumCache(series.Values))));

        return ser;
    }

    private static XElement StrCache(IReadOnlyList<string> labels)
    {
        XElement cache = new(C + "strCache", new XElement(C + "ptCount", new XAttribute("val", labels.Count)));
        for (int i = 0; i < labels.Count; i++)
        {
            cache.Add(new XElement(C + "pt", new XAttribute("idx", i), new XElement(C + "v", labels[i] ?? string.Empty)));
        }

        return cache;
    }

    private static XElement NumCache(IReadOnlyList<double> values)
    {
        XElement cache = new(C + "numCache", new XElement(C + "formatCode", "General"), new XElement(C + "ptCount", new XAttribute("val", values.Count)));
        for (int i = 0; i < values.Count; i++)
        {
            double v = values[i];
            string text = double.IsNaN(v) || double.IsInfinity(v) ? "0" : v.ToString(CultureInfo.InvariantCulture);
            cache.Add(new XElement(C + "pt", new XAttribute("idx", i), new XElement(C + "v", text)));
        }

        return cache;
    }

    private static XElement BuildCategoryAxis() => new(
        C + "catAx",
        new XElement(C + "axId", new XAttribute("val", 1)),
        new XElement(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax"))),
        new XElement(C + "delete", new XAttribute("val", 0)),
        new XElement(C + "axPos", new XAttribute("val", "b")),
        new XElement(C + "crossAx", new XAttribute("val", 2)));

    private static XElement BuildValueAxis(double? minimum)
    {
        XElement scaling = new(C + "scaling", new XElement(C + "orientation", new XAttribute("val", "minMax")));
        if (minimum is { } min)
        {
            scaling.Add(new XElement(C + "min", new XAttribute("val", min.ToString(CultureInfo.InvariantCulture))));
        }

        return new XElement(
            C + "valAx",
            new XElement(C + "axId", new XAttribute("val", 2)),
            scaling,
            new XElement(C + "delete", new XAttribute("val", 0)),
            new XElement(C + "axPos", new XAttribute("val", "l")),
            new XElement(C + "crossAx", new XAttribute("val", 1)));
    }

    private static (string Element, string? BarDir, string? Grouping, bool HasAxes) ChartElement(ChartKind kind) => kind switch
    {
        ChartKind.Bar => ("barChart", "col", "clustered", true),
        ChartKind.StackedBar => ("barChart", "col", "stacked", true),
        ChartKind.Line or ChartKind.Spline or ChartKind.StepLine => ("lineChart", null, "standard", true),
        ChartKind.Area => ("areaChart", null, "standard", true),
        ChartKind.StackedArea => ("areaChart", null, "stacked", true),
        ChartKind.Pie => ("pieChart", null, null, false),
        ChartKind.Donut => ("doughnutChart", null, null, false),
        ChartKind.Radar => ("radarChart", null, null, true),
        ChartKind.Scatter => ("lineChart", null, "standard", true), // scatter approximated as a line series
        _ => ("barChart", "col", "clustered", true),
    };

    // ----- package part helpers -----

    private static XDocument BuildRels(IEnumerable<(string Id, string Type, string Target)> rels)
    {
        XElement root = new(Rels + "Relationships");
        foreach ((string id, string type, string target) in rels)
        {
            root.Add(new XElement(
                Rels + "Relationship",
                new XAttribute("Id", id),
                new XAttribute("Type", type),
                new XAttribute("Target", target)));
        }

        return new XDocument(root);
    }

    private static XDocument AddRelationship(byte[]? existing, string id, string type, string target)
    {
        XDocument doc = existing is not null ? Parse(existing) : new XDocument(new XElement(Rels + "Relationships"));
        doc.Root!.Add(new XElement(
            Rels + "Relationship",
            new XAttribute("Id", id),
            new XAttribute("Type", type),
            new XAttribute("Target", target)));
        return doc;
    }

    private static XDocument AddDrawingElement(byte[] worksheetXml, string drawingRelId)
    {
        XDocument doc = Parse(worksheetXml);
        XElement root = doc.Root!;

        // Ensure the relationships namespace prefix 'r' is available for the r:id attribute.
        if (root.Attribute(XNamespace.Xmlns + "r") is null)
        {
            root.Add(new XAttribute(XNamespace.Xmlns + "r", R.NamespaceName));
        }

        root.Elements(Main + "drawing").Remove();
        XElement drawing = new(Main + "drawing", new XAttribute(R + "id", drawingRelId));
        // OOXML places drawings before tableParts (and the other worksheet tail elements).
        // Appending after a preserved table definition produces a schema-invalid worksheet.
        XElement? next = root.Elements().FirstOrDefault(element => element.Name.Namespace == Main &&
            element.Name.LocalName is "legacyDrawing" or "legacyDrawingHF" or "picture" or "oleObjects" or
                "controls" or "webPublishItems" or "tableParts" or "extLst");
        if (next is null) root.Add(drawing);
        else next.AddBeforeSelf(drawing);
        return doc;
    }

    private static XDocument AddContentTypes(byte[] contentTypesXml, IReadOnlyList<(string PartName, string ContentType)> overrides)
    {
        XDocument doc = Parse(contentTypesXml);
        XElement root = doc.Root!;
        foreach ((string partName, string contentType) in overrides)
        {
            bool present = root.Elements(Ct + "Override").Any(e => string.Equals(e.Attribute("PartName")?.Value, partName, StringComparison.OrdinalIgnoreCase));
            if (!present)
            {
                root.Add(new XElement(Ct + "Override", new XAttribute("PartName", partName), new XAttribute("ContentType", contentType)));
            }
        }

        return doc;
    }

    private static XDocument Parse(byte[] xml)
    {
        using MemoryStream ms = new(xml);
        return XDocument.Load(ms);
    }

    private static byte[] ToBytes(XDocument doc)
    {
        using MemoryStream ms = new();
        // OOXML parts are UTF-8 with an XML declaration.
        using (System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(ms, new System.Xml.XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false }))
        {
            doc.Save(writer);
        }

        return ms.ToArray();
    }
}
