namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Models;

public static partial class XlsxFile
{
    private static void ReadAdditionalConditionalFormatting(ZipArchive archive, string part, IReadOnlyList<Color?> theme,
        List<SheetCellValueRule> rules, CancellationToken cancellationToken)
    {
        XElement[] dxfs = archive.GetEntry("xl/styles.xml") is not null
            ? ReadXmlPart(archive, "xl/styles.xml").Descendants(Ssml + "dxfs").Elements().ToArray()
            : Array.Empty<XElement>();
        using Stream input = archive.GetEntry(part)!.Open();
        using XmlReader reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheetData") { reader.Skip(); continue; }
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "conditionalFormatting") { reader.Read(); continue; }
            XElement block = (XElement)XNode.ReadFrom(reader);
            foreach (XElement rule in block.Elements(Ssml + "cfRule"))
            {
                string? type = rule.Attribute("type")?.Value, formula = rule.Element(Ssml + "formula")?.Value;
                bool text = type == "cellIs" && formula is { Length: >= 2 } && formula.StartsWith('"') && formula.EndsWith('"');
                if (!text && type is not ("expression" or "iconSet")) continue;
                int dxfId = (int?)rule.Attribute("dxfId") ?? -1;
                XElement? dxf = dxfId >= 0 && dxfId < dxfs.Length ? dxfs[dxfId] : null;
                XElement? icon = rule.Element(Ssml + "iconSet");
                foreach (string reference in (block.Attribute("sqref")?.Value ?? string.Empty).Split(' ', StringSplitOptions.RemoveEmptyEntries))
                {
                    if (!CellRange.TryParse(reference, out CellRange range)) continue;
                    TryComparison(rule.Attribute("operator")?.Value, out CellValueComparison comparison);
                    rules.Add(new SheetCellValueRule
                    {
                        Range = range, Comparison = comparison, Threshold = 0, Priority = (int?)rule.Attribute("priority") ?? int.MaxValue,
                        Expression = type == "expression" ? formula : null,
                        TextThreshold = text ? formula![1..^1].Replace("\"\"", "\"") : null,
                        Foreground = PreviewColor(dxf?.Element(Ssml + "font")?.Element(Ssml + "color"), theme),
                        Background = PreviewColor(dxf?.Element(Ssml + "fill")?.Element(Ssml + "patternFill")?.Element(Ssml + "fgColor"), theme),
                        Bold = dxf?.Element(Ssml + "font")?.Element(Ssml + "b") is { } bold ? bold.Attribute("val")?.Value is not ("0" or "false") : null,
                        IconSet = icon?.Attribute("iconSet")?.Value,
                        IconThresholds = icon?.Elements(Ssml + "cfvo").Select(e => double.TryParse(e.Attribute("val")?.Value, CultureInfo.InvariantCulture, out double n) ? n : 0).ToArray() ?? Array.Empty<double>(),
                        ShowValue = icon?.Attribute("showValue")?.Value is not ("0" or "false"),
                        ReverseIcons = icon?.Attribute("reverse")?.Value is "1" or "true",
                    });
                }
            }
        }
    }

    private static void WriteAdditionalConditionalFormatting(XDocument worksheet, XDocument styles, SheetData data)
    {
        XElement? dxfs = styles.Root!.Element(Ssml + "dxfs");
        if (dxfs is null) { dxfs = new XElement(Ssml + "dxfs", new XAttribute("count", 0)); styles.Root.Add(dxfs); }
        foreach (SheetCellValueRule rule in data.CellValueRules.Where(r => r.Expression is not null || r.TextThreshold is not null || r.IconSet is not null))
        {
            XElement element = new(Ssml + "cfRule", new XAttribute("type", rule.IconSet is not null ? "iconSet" : rule.Expression is not null ? "expression" : "cellIs"), new XAttribute("priority", Math.Max(1, rule.Priority)));
            if (rule.IconSet is { } iconSet)
                element.Add(new XElement(Ssml + "iconSet", new XAttribute("iconSet", iconSet), new XAttribute("showValue", rule.ShowValue ? "1" : "0"), new XAttribute("reverse", rule.ReverseIcons ? "1" : "0"),
                    rule.IconThresholds.Select(n => new XElement(Ssml + "cfvo", new XAttribute("type", "num"), new XAttribute("val", n)))));
            else
            {
                XElement dxf = new(Ssml + "dxf"), font = new(Ssml + "font");
                if (rule.Bold is { } bold) font.Add(new XElement(Ssml + "b", new XAttribute("val", bold ? "1" : "0")));
                if (rule.Foreground is { } fg) font.Add(new XElement(Ssml + "color", new XAttribute("rgb", $"FF{fg.R:X2}{fg.G:X2}{fg.B:X2}")));
                if (font.HasElements) dxf.Add(font);
                if (rule.Background is { } bg) dxf.Add(new XElement(Ssml + "fill", new XElement(Ssml + "patternFill", new XAttribute("patternType", "solid"), new XElement(Ssml + "fgColor", new XAttribute("rgb", $"FF{bg.R:X2}{bg.G:X2}{bg.B:X2}")))));
                int id = dxfs.Elements().Count(); dxfs.Add(dxf); dxfs.SetAttributeValue("count", id + 1); element.SetAttributeValue("dxfId", id);
                if (rule.TextThreshold is { } text) element.SetAttributeValue("operator", rule.Comparison == CellValueComparison.NotEqual ? "notEqual" : "equal");
                element.Add(new XElement(Ssml + "formula", rule.Expression ?? "\"" + rule.TextThreshold!.Replace("\"", "\"\"") + "\""));
            }
            XElement block = new(Ssml + "conditionalFormatting", new XAttribute("sqref", rule.Range.ToA1()), element);
            XElement? before = worksheet.Root!.Elements().FirstOrDefault(e => e.Name.LocalName is "dataValidations" or "hyperlinks" or "printOptions" or "pageMargins" or "pageSetup" or "headerFooter" or "drawing" or "tableParts" or "extLst");
            if (before is null) worksheet.Root.Add(block); else before.AddBeforeSelf(block);
        }
    }
}
