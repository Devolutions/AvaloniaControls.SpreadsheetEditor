namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Avalonia.Layout;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.Formulas;

public static partial class XlsxFile
{
    private static readonly IReadOnlyDictionary<int, string> EmptyFormulaSharedStrings = new Dictionary<int, string>();

    private sealed record NativeSheetPresentation(
        IReadOnlyList<int> HiddenColumns, IReadOnlyList<SheetOutlineGroup> RowGroups,
        IReadOnlyList<SheetOutlineGroup> ColumnGroups, string? ProtectionXml, IReadOnlyList<MergeRange> Merges,
        IReadOnlyList<PreviewColumn> Columns);

    private static List<(CellAppearance Appearance, string? Horizontal)> ReadAppearanceStyles(ZipArchive archive)
    {
        static List<(CellAppearance, string?)> Defaults() => new() { (new CellAppearance(), null) };
        if (archive.GetEntry("xl/styles.xml") is null) return Defaults();
        XDocument styles = ReadXmlPart(archive, "xl/styles.xml");
        XElement[] fonts = styles.Root?.Element(Ssml + "fonts")?.Elements().ToArray() ?? Array.Empty<XElement>();
        XElement[] fills = styles.Root?.Element(Ssml + "fills")?.Elements().ToArray() ?? Array.Empty<XElement>();
        XElement[] borders = styles.Root?.Element(Ssml + "borders")?.Elements().ToArray() ?? Array.Empty<XElement>();
        IReadOnlyList<Color?> theme = ReadThemeColors(archive);
        List<(CellAppearance, string?)> result = new();
        foreach (XElement xf in styles.Root?.Element(Ssml + "cellXfs")?.Elements() ?? Enumerable.Empty<XElement>())
        {
            XElement? Part(XElement[] parts, string attribute) =>
                int.TryParse(xf.Attribute(attribute)?.Value, out int i) && i >= 0 && i < parts.Length ? parts[i] : null;
            XElement? font = Part(fonts, "fontId"), fill = Part(fills, "fillId"), border = Part(borders, "borderId");
            string? script = font?.Element(Ssml + "vertAlign")?.Attribute("val")?.Value;
            XElement? alignment = xf.Element(Ssml + "alignment");
            CellEdge? Edge(string name)
            {
                XElement? edge = border?.Element(Ssml + name);
                return edge?.Attribute("style") is { } style && style.Value != "none"
                    ? new CellEdge(style.Value, PreviewColor(edge.Element(Ssml + "color"), theme) ?? Colors.Black) : null;
            }
            XElement? pattern = fill?.Element(Ssml + "patternFill"), gradient = fill?.Element(Ssml + "gradientFill");
            CellFill? cellFill = gradient is not null ? new CellFill
            {
                Degree = (double?)gradient.Attribute("degree") ?? 0,
                Stops = gradient.Elements(Ssml + "stop").Select(stop => new CellGradientStop(
                    (double?)stop.Attribute("position") ?? 0, PreviewColor(stop.Element(Ssml + "color"), theme) ?? Colors.White)).ToArray(),
            } : pattern?.Attribute("patternType")?.Value is { } patternType && patternType is not "none" and not "solid"
                ? new CellFill { Pattern = patternType, Foreground = PreviewColor(pattern.Element(Ssml + "fgColor"), theme), Background = PreviewColor(pattern.Element(Ssml + "bgColor"), theme) }
                : null;
            result.Add((new CellAppearance
            {
                VerticalAlignment = alignment?.Attribute("vertical")?.Value switch { "top" => VerticalAlignment.Top, "center" => VerticalAlignment.Center, _ => VerticalAlignment.Bottom },
                Rotation = (int?)alignment?.Attribute("textRotation") ?? 0,
                Indent = (int?)alignment?.Attribute("indent") ?? 0,
                ShrinkToFit = alignment?.Attribute("shrinkToFit")?.Value is "1" or "true",
                DoubleUnderline = font?.Element(Ssml + "u")?.Attribute("val")?.Value is "double" or "doubleAccounting",
                Script = script is "superscript" or "subscript" ? script : null,
                Locked = xf.Element(Ssml + "protection")?.Attribute("locked")?.Value is not ("0" or "false"),
                Left = Edge("left"), Right = Edge("right"), Top = Edge("top"), Bottom = Edge("bottom"), Fill = cellFill,
            }, alignment?.Attribute("horizontal")?.Value));
        }
        return result.Count > 0 ? result : Defaults();
    }

    // Content-only used ranges omit blank cells carrying borders. Scan explicit cell records so
    // formatting an entire row/column does not materialize Excel's whole worksheet on import.
    private static (int Rows, int Columns) ReadBorderExtent(ZipArchive archive, string part, CancellationToken cancellationToken)
    {
        List<(CellAppearance Appearance, string? Horizontal)> styles = ReadAppearanceStyles(archive);
        bool[] bordered = styles.Select(style => style.Appearance.Left is not null || style.Appearance.Right is not null
            || style.Appearance.Top is not null || style.Appearance.Bottom is not null).ToArray();
        if (!bordered.Any(static value => value)) return (0, 0);
        int rows = 0, columns = 0;
        using Stream input = archive.GetEntry(part)!.Open();
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
            {
                if (int.TryParse(reader.GetAttribute("s"), out int style) && style >= 0 && style < bordered.Length && bordered[style]
                    && CellRange.TryParse(reader.GetAttribute("r"), out CellRange range))
                {
                    rows = Math.Max(rows, range.Row + range.RowSpan);
                    columns = Math.Max(columns, range.Column + range.ColumnSpan);
                }
                reader.Skip();
                continue;
            }
            reader.Read();
        }
        return (rows, columns);
    }

    private static HashSet<(int Row, int Column)> ReadAutomaticTableFonts(ZipArchive archive, string part,
        IReadOnlyList<SheetTable> tables, CancellationToken cancellationToken)
    {
        HashSet<(int Row, int Column)> result = new();
        if (tables.Count == 0)
        {
            return result;
        }
        bool[] automaticStyles = ReadPreviewStyles(archive).Select(static style => style.AutomaticForeground).ToArray();
        if (!automaticStyles.Any(static automatic => automatic))
        {
            return result;
        }
        // ClosedXML resolves table fonts before returning cell.Style, so retain explicit Automatic separately.
        using Stream input = archive.GetEntry(part)!.Open();
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
            {
                int style = int.TryParse(reader.GetAttribute("s"), out int parsed) ? parsed : 0;
                if (style >= 0 && style < automaticStyles.Length && automaticStyles[style]
                    && TryParseCellAddress(reader.GetAttribute("r"), out int row, out int column)
                    && tables.Any(table => row >= table.Range.Row && row < table.Range.Row + table.Range.RowSpan
                        && column >= table.Range.Column && column < table.Range.Column + table.Range.ColumnSpan))
                {
                    result.Add((row, column));
                }
                reader.Skip();
                continue;
            }
            reader.Read();
        }
        return result;
    }

    // Only cell attributes and outline metadata are read; large cell contents never become a DOM.
    private static NativeSheetPresentation ReadNativePresentation(ZipArchive archive, string part,
        List<List<CellData>> rows, CancellationToken cancellationToken)
    {
        List<(CellAppearance Appearance, string? Horizontal)> styles = ReadAppearanceStyles(archive);
        List<int> hiddenColumns = new();
        List<PreviewColumn> columns = new();
        List<MergeRange> merges = new();
        IReadOnlyDictionary<string, string> hyperlinkTargets = ReadHyperlinkRelationships(archive, part);
        Dictionary<int, (int Level, bool Hidden)> rowLevels = new(), columnLevels = new();
        bool summaryBelow = true, summaryRight = true;
        string? protection = null;
        using Stream input = archive.GetEntry(part)!.Open();
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element) { reader.Read(); continue; }
            switch (reader.LocalName)
            {
                case "outlinePr":
                    summaryBelow = reader.GetAttribute("summaryBelow") is not ("0" or "false");
                    summaryRight = reader.GetAttribute("summaryRight") is not ("0" or "false");
                    break;
                case "sheetProtection":
                    protection = ((XElement)XNode.ReadFrom(reader)).ToString(SaveOptions.DisableFormatting);
                    continue;
                case "mergeCell":
                    if (CellRange.TryParse(reader.GetAttribute("ref"), out CellRange range))
                        merges.Add(new MergeRange(range.Row, range.Column, range.RowSpan, range.ColumnSpan));
                    break;
                case "hyperlink":
                    if (CellRange.TryParse(reader.GetAttribute("ref"), out CellRange linkRange))
                    {
                        string? id = reader.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        string? target = id is not null && hyperlinkTargets.TryGetValue(id, out string? external) ? external : null;
                        if (reader.GetAttribute("location") is { Length: > 0 } location)
                            target = target is null ? "#" + location : target + "#" + location;
                        if (target is not null)
                        {
                            for (int linkRow = linkRange.Row; linkRow < rows.Count && linkRow < linkRange.Row + linkRange.RowSpan; linkRow++)
                                for (int linkColumn = linkRange.Column; linkColumn < rows[linkRow].Count && linkColumn < linkRange.Column + linkRange.ColumnSpan; linkColumn++)
                                    rows[linkRow][linkColumn] = rows[linkRow][linkColumn] with { Hyperlink = target };
                        }
                    }
                    break;
                case "col":
                    int min = int.TryParse(reader.GetAttribute("min"), out int a) ? a - 1 : 0;
                    int max = int.TryParse(reader.GetAttribute("max"), out int b) ? Math.Min(b, 16384) : min + 1;
                    int level = int.TryParse(reader.GetAttribute("outlineLevel"), out int l) ? l : 0;
                    bool hidden = reader.GetAttribute("hidden") is "1" or "true";
                    double? width = double.TryParse(reader.GetAttribute("width"), System.Globalization.NumberStyles.Float,
                        System.Globalization.CultureInfo.InvariantCulture, out double parsedWidth) ? parsedWidth : null;
                    columns.Add(new PreviewColumn(min + 1, max, width, hidden));
                    for (int c = min; c < max; c++) { if (hidden) hiddenColumns.Add(c); if (level > 0) columnLevels[c] = (level, hidden); }
                    break;
                case "row":
                    if (int.TryParse(reader.GetAttribute("r"), out int row) && int.TryParse(reader.GetAttribute("outlineLevel"), out int rowLevel) && rowLevel > 0)
                        rowLevels[row - 1] = (rowLevel, reader.GetAttribute("hidden") is "1" or "true");
                    break;
                case "c":
                    bool consumed = false;
                    if (TryParseCellAddress(reader.GetAttribute("r"), out int r, out int c2) && r < rows.Count && c2 < rows[r].Count)
                    {
                        int index = int.TryParse(reader.GetAttribute("s"), out int s) ? s : 0;
                        if (index >= 0 && index < styles.Count)
                        {
                            (CellAppearance appearance, string? horizontal) = styles[index];
                            CellData cell = rows[r][c2];
                            string? type = reader.GetAttribute("t");
                            TextAlignment align = horizontal switch
                            {
                                "center" or "centerContinuous" => TextAlignment.Center, "right" => TextAlignment.Right,
                                "left" => TextAlignment.Left, _ => type == "b" ? TextAlignment.Center :
                                    type is null or "n" or "d" ? TextAlignment.Right : TextAlignment.Left,
                            };
                            rows[r][c2] = cell with { Appearance = appearance, Alignment = align, Background = appearance.Fill is null ? cell.Background : null };
                        }
                        // A missing cache is different from a valid, deliberately empty string.
                        // ClosedXML represents both as Blank, so consult the native cell type.
                        CellData styled = rows[r][c2];
                        if (FormulaEngine.IsFormula(styled.Text) && styled.FormulaValue is null)
                        {
                            XElement rawCell = (XElement)XNode.ReadFrom(reader);
                            consumed = true;
                            if (rawCell.Element(Ssml + "f") is not null)
                            {
                                CellValue? cached = ReadPreviewFormulaValue(rawCell, EmptyFormulaSharedStrings);
                                rows[r][c2] = styled with
                                {
                                    FormulaValue = cached,
                                    FormulaNumericValue = cached is { Kind: CellValueKind.Number } n ? n.Number : null,
                                    FormulaResult = cached is { Kind: CellValueKind.Number } number ? NumberFormatter.Format(number.Number, styled.NumberFormat) : cached?.Display(),
                                };
                            }
                        }
                    }
                    if (!consumed) reader.Skip();
                    continue;
            }
            reader.Read();
        }
        static IReadOnlyList<SheetOutlineGroup> Groups(Dictionary<int, (int Level, bool Hidden)> levels, bool after)
        {
            List<SheetOutlineGroup> groups = new();
            for (int level = 1; level <= 7; level++)
            {
                int[] indexes = levels.Where(p => p.Value.Level >= level).Select(p => p.Key).Order().ToArray();
                for (int i = 0; i < indexes.Length; i++)
                {
                    int start = indexes[i], end = start;
                    while (i + 1 < indexes.Length && indexes[i + 1] == end + 1) end = indexes[++i];
                    groups.Add(new SheetOutlineGroup(start, end - start + 1, level,
                        Enumerable.Range(start, end - start + 1).All(n => levels[n].Hidden), after));
                }
            }
            return groups;
        }
        return new(hiddenColumns, Groups(rowLevels, summaryBelow), Groups(columnLevels, summaryRight), protection, merges, columns);
    }

    private static void RestorePresentation(MemoryStream package, IReadOnlyList<XlsxSheetWriteData> sheets)
    {
        package.Position = 0;
        using ZipArchive archive = new(package, ZipArchiveMode.Update, leaveOpen: true);
        XDocument styles = ReadXmlPart(archive, "xl/styles.xml");
        XElement fonts = styles.Root!.Element(Ssml + "fonts")!, fills = styles.Root.Element(Ssml + "fills")!,
            borders = styles.Root.Element(Ssml + "borders")!, xfs = styles.Root.Element(Ssml + "cellXfs")!;
        // Indexed 64 is ClosedXML's automatic colour; use Excel's explicit automatic font marker.
        foreach (XElement color in fonts.Elements().Elements(Ssml + "color"))
        {
            if (color.Attribute("indexed")?.Value == "64")
            {
                color.ReplaceWith(new XElement(Ssml + "color", new XAttribute("auto", 1)));
            }
        }
        List<XElement> originalFonts = fonts.Elements().ToList(), originalXfs = xfs.Elements().ToList();
        Dictionary<(int Base, CellAppearance Appearance), int> styleIds = new();
        Dictionary<string, int> fillIds = new(), borderIds = new(), fontIds = new();
        static int Intern(XElement collection, XElement value, Dictionary<string, int> cache)
        {
            string key = value.ToString(SaveOptions.DisableFormatting);
            if (cache.TryGetValue(key, out int id)) return id;
            id = collection.Elements().Count();
            collection.Add(value);
            cache[key] = id;
            collection.SetAttributeValue("count", id + 1);
            return id;
        }
        XElement ColorElement(Color color) => new(Ssml + "color", new XAttribute("rgb", $"FF{color.R:X2}{color.G:X2}{color.B:X2}"));
        int StyleId(int original, CellAppearance appearance)
        {
            if (styleIds.TryGetValue((original, appearance), out int id)) return id;
            XElement xf = new(originalXfs[original]);
            XElement alignment = xf.Element(Ssml + "alignment") ?? new XElement(Ssml + "alignment");
            if (alignment.Parent is null) xf.Add(alignment);
            alignment.SetAttributeValue("vertical", appearance.VerticalAlignment switch { VerticalAlignment.Top => "top", VerticalAlignment.Center => "center", _ => "bottom" });
            alignment.SetAttributeValue("textRotation", appearance.Rotation);
            alignment.SetAttributeValue("indent", appearance.Indent);
            alignment.SetAttributeValue("shrinkToFit", appearance.ShrinkToFit ? "1" : "0");
            xf.SetAttributeValue("applyAlignment", "1");
            xf.Elements(Ssml + "protection").Remove();
            xf.Add(new XElement(Ssml + "protection", new XAttribute("locked", appearance.Locked ? "1" : "0")));
            xf.SetAttributeValue("applyProtection", "1");
            if (appearance.Script is not null || appearance.DoubleUnderline)
            {
                int oldFont = (int?)xf.Attribute("fontId") ?? 0;
                XElement font = new(originalFonts[oldFont]);
                if (appearance.Script is not null)
                {
                    font.Elements(Ssml + "vertAlign").Remove();
                    font.Add(new XElement(Ssml + "vertAlign", new XAttribute("val", appearance.Script)));
                }
                if (appearance.DoubleUnderline)
                {
                    font.Elements(Ssml + "u").Remove();
                    font.Add(new XElement(Ssml + "u", new XAttribute("val", "double")));
                }
                string[] fontOrder = { "b", "i", "strike", "condense", "extend", "outline", "shadow", "u", "vertAlign", "sz", "color", "name", "family", "charset", "scheme" };
                font.ReplaceNodes(font.Elements().OrderBy(e => Array.IndexOf(fontOrder, e.Name.LocalName)).ToArray());
                xf.SetAttributeValue("fontId", Intern(fonts, font, fontIds));
                xf.SetAttributeValue("applyFont", "1");
            }
            if (appearance.Fill is { } fill)
            {
                XElement content;
                if (fill.Stops.Count > 0)
                    content = new XElement(Ssml + "gradientFill", new XAttribute("degree", fill.Degree), fill.Stops.Select(stop =>
                        new XElement(Ssml + "stop", new XAttribute("position", stop.Position), ColorElement(stop.Color))));
                else
                {
                    content = new XElement(Ssml + "patternFill", new XAttribute("patternType", fill.Pattern ?? "none"));
                    if (fill.Foreground is { } fg) { XElement color = ColorElement(fg); color.Name = Ssml + "fgColor"; content.Add(color); }
                    if (fill.Background is { } bg) { XElement color = ColorElement(bg); color.Name = Ssml + "bgColor"; content.Add(color); }
                }
                xf.SetAttributeValue("fillId", Intern(fills, new XElement(Ssml + "fill", content), fillIds));
                xf.SetAttributeValue("applyFill", "1");
            }
            if (appearance.Left is not null || appearance.Right is not null || appearance.Top is not null || appearance.Bottom is not null)
            {
                XElement border = new(Ssml + "border");
                foreach ((string name, CellEdge? edge) in new[] { ("left", appearance.Left), ("right", appearance.Right), ("top", appearance.Top), ("bottom", appearance.Bottom) })
                    border.Add(edge is null ? new XElement(Ssml + name) : new XElement(Ssml + name, new XAttribute("style", edge.Style), ColorElement(edge.Color)));
                border.Add(new XElement(Ssml + "diagonal"));
                xf.SetAttributeValue("borderId", Intern(borders, border, borderIds));
                xf.SetAttributeValue("applyBorder", "1");
            }
            // alignment precedes protection in CT_Xf.
            xf.Elements(Ssml + "alignment").Remove();
            xf.AddFirst(alignment);
            id = xfs.Elements().Count();
            xfs.Add(xf);
            xfs.SetAttributeValue("count", id + 1);
            styleIds[(original, appearance)] = id;
            return id;
        }
        static void Write(ZipArchive archive, string path, XDocument document)
        {
            archive.GetEntry(path)?.Delete();
            using Stream output = archive.CreateEntry(path, CompressionLevel.Optimal).Open();
            using XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings
            {
                Encoding = new System.Text.UTF8Encoding(false), NewLineHandling = NewLineHandling.Entitize,
            });
            document.Save(writer);
        }
        XDocument workbook = ReadXmlPart(archive, "xl/workbook.xml");
        XElement[] sheetElements = workbook.Descendants(Ssml + "sheet").ToArray();
        for (int i = 0; i < sheets.Count; i++)
        {
            SheetData data = sheets[i].Data;
            sheetElements[i].SetAttributeValue("state", data.Visibility is "hidden" or "veryhidden" or "veryHidden" ? data.Visibility.Equals("hidden", StringComparison.Ordinal) ? "hidden" : "veryHidden" : "visible");
            string part = WorksheetPartPath(archive, i);
            XDocument document = ReadXmlPart(archive, part);
            WriteAdditionalConditionalFormatting(document, styles, data);
            foreach (XElement cell in document.Descendants(Ssml + "c"))
                if (TryParseCellAddress(cell.Attribute("r")?.Value, out int row, out int column) && row < data.Cells.Count && column < data.Cells[row].Count && data.Cells[row][column].Appearance is { } appearance)
                    cell.SetAttributeValue("s", StyleId((int?)cell.Attribute("s") ?? 0, appearance));
            if (data.ProtectionXml is not null)
            {
                document.Root!.Elements(Ssml + "sheetProtection").Remove();
                document.Root.Element(Ssml + "sheetData")!.AddAfterSelf(XElement.Parse(data.ProtectionXml));
            }
            foreach (int column in data.HiddenColumns)
            {
                XElement? cols = document.Root!.Element(Ssml + "cols");
                if (cols is null) { cols = new XElement(Ssml + "cols"); document.Root.Element(Ssml + "sheetData")!.AddBeforeSelf(cols); }
                // ClosedXML creates one width record per column; split a grouped record when needed.
                XElement? existing = cols.Elements().FirstOrDefault(c => (int?)c.Attribute("min") <= column + 1 && (int?)c.Attribute("max") >= column + 1);
                if (existing is not null)
                {
                    int first = (int)existing.Attribute("min")!, last = (int)existing.Attribute("max")!;
                    if (first < column + 1) { XElement before = new(existing); before.SetAttributeValue("max", column); existing.AddBeforeSelf(before); }
                    if (last > column + 1) { XElement after = new(existing); after.SetAttributeValue("min", column + 2); existing.AddAfterSelf(after); }
                    existing.SetAttributeValue("min", column + 1); existing.SetAttributeValue("max", column + 1); existing.SetAttributeValue("hidden", "1");
                }
                else cols.Add(new XElement(Ssml + "col", new XAttribute("min", column + 1), new XAttribute("max", column + 1), new XAttribute("width", 13), new XAttribute("customWidth", "1"), new XAttribute("hidden", "1")));
            }
            XElement? sheetPr = document.Root!.Element(Ssml + "sheetPr");
            if (data.RowGroups.Count > 0 || data.ColumnGroups.Count > 0)
            {
                if (sheetPr is null) { sheetPr = new XElement(Ssml + "sheetPr"); document.Root.AddFirst(sheetPr); }
                sheetPr.Elements(Ssml + "outlinePr").Remove();
                sheetPr.Add(new XElement(Ssml + "outlinePr", new XAttribute("summaryBelow", data.RowGroups.FirstOrDefault()?.SummaryAfter == false ? "0" : "1"), new XAttribute("summaryRight", data.ColumnGroups.FirstOrDefault()?.SummaryAfter == false ? "0" : "1")));
                foreach (XElement row in document.Descendants(Ssml + "row"))
                {
                    int r = (int)row.Attribute("r")! - 1;
                    int level = data.RowGroups.Where(g => r >= g.Start && r < g.Start + g.Count).Select(g => g.Level).DefaultIfEmpty().Max();
                    if (level > 0) row.SetAttributeValue("outlineLevel", level);
                }
                // Column groups use per-column records, preserving hidden state and widths.
                foreach (XElement col in document.Root.Element(Ssml + "cols")?.Elements().ToArray() ?? Array.Empty<XElement>())
                {
                    int first = (int)col.Attribute("min")!, last = (int)col.Attribute("max")!;
                    for (int c = first; c <= last; c++)
                    {
                        XElement split = new(col); split.SetAttributeValue("min", c); split.SetAttributeValue("max", c);
                        int level = data.ColumnGroups.Where(g => c - 1 >= g.Start && c - 1 < g.Start + g.Count).Select(g => g.Level).DefaultIfEmpty().Max();
                        if (level > 0) split.SetAttributeValue("outlineLevel", level);
                        col.AddBeforeSelf(split);
                    }
                    col.Remove();
                }
            }
            Write(archive, part, document);
        }
        int firstVisible = Array.FindIndex(sheetElements, sheet => sheet.Attribute("state")?.Value is null or "visible");
        if (firstVisible >= 0)
        {
            foreach (XElement view in workbook.Descendants(Ssml + "workbookView"))
            {
                int active = (int?)view.Attribute("activeTab") ?? 0;
                if (active < 0 || active >= sheetElements.Length || sheetElements[active].Attribute("state")?.Value is "hidden" or "veryHidden")
                    view.SetAttributeValue("activeTab", firstVisible);
            }
        }
        Write(archive, "xl/styles.xml", styles);
        Write(archive, "xl/workbook.xml", workbook);
    }
}
