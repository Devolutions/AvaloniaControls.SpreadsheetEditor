namespace SpreadsheetEditor.Controls.Clipboard;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

using Avalonia.Layout;
using Avalonia.Media;

using HtmlAgilityPack;

using global::SpreadsheetEditor.Controls.Models;

// Excel supplies a small HTML document, including class rules in its head, on the clipboard.
// Parse it as data only: no browser, external style sheets, images or scripts are involved.
internal sealed record ClipboardTable(CellData[][] Cells, IReadOnlyList<MergeRange> Merges)
{
    private const int MaxHtmlCells = 100_000;
    private static readonly Regex Rules = new(@"([^{}]+)\{([^{}]*)\}", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
    private static readonly Regex SimpleSelector = new(@"^(?<tag>[a-z][a-z0-9]*)?(?<classes>(?:\.[\w-]+)*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    internal static ClipboardTable? FromHtml(string html)
    {
        HtmlDocument document = new() { OptionMaxNestedChildNodes = 256 };
        document.LoadHtml(html.TrimEnd('\0'));
        HtmlNode? table = document.DocumentNode.Descendants("table").FirstOrDefault();
        if (table is null) return null;

        List<StyleRule> rules = new();
        foreach (HtmlNode style in document.DocumentNode.Descendants("style"))
        {
            string css = Regex.Replace(style.InnerText.Replace("<!--", "").Replace("-->", ""), @"/\*[\s\S]*?\*/", "", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
            foreach (Match match in Rules.Matches(css))
            {
                foreach (string selector in match.Groups[1].Value.Split(','))
                {
                    Match simple = SimpleSelector.Match(selector.Trim());
                    if (!simple.Success) continue;
                    string[] classes = simple.Groups["classes"].Value.Split('.', StringSplitOptions.RemoveEmptyEntries);
                    string tag = simple.Groups["tag"].Value;
                    rules.Add(new StyleRule(tag, classes, Declarations(match.Groups[2].Value), classes.Length * 10 + (tag.Length > 0 ? 1 : 0)));
                }
            }
        }
        rules = rules.OrderBy(rule => rule.Specificity).ToList();

        Dictionary<(int Row, int Column), CellData> cells = new();
        List<MergeRange> merges = new();
        int rowIndex = 0, width = 0, height = 0;
        foreach (HtmlNode row in table.Descendants("tr").Where(row => row.Ancestors("table").First() == table))
        {
            int column = 0;
            foreach (HtmlNode cell in row.ChildNodes.Where(node => node.Name is "td" or "th"))
            {
                while (cells.ContainsKey((rowIndex, column))) column++;
                int rows = Math.Max(1, cell.GetAttributeValue("rowspan", 1));
                int columns = Math.Max(1, cell.GetAttributeValue("colspan", 1));
                if (rows > MaxHtmlCells || columns > MaxHtmlCells ||
                    (long)rows * columns + cells.Count > MaxHtmlCells || (long)(rowIndex + rows) * (column + columns) > MaxHtmlCells)
                    return null;
                cells[(rowIndex, column)] = ReadCell(cell, rules);
                for (int r = rowIndex; r < rowIndex + rows; r++)
                {
                    for (int c = column; c < column + columns; c++)
                    {
                        if (r != rowIndex || c != column) cells[(r, c)] = CellData.Plain(string.Empty);
                    }
                }
                if (rows > 1 || columns > 1) merges.Add(new MergeRange(rowIndex, column, rows, columns));
                width = Math.Max(width, column + columns);
                height = Math.Max(height, rowIndex + rows);
                column += columns;
            }
            rowIndex++;
            if ((long)rowIndex * Math.Max(width, 1) > MaxHtmlCells) return null;
        }
        if (width == 0 || height == 0 || (long)width * height > MaxHtmlCells) return null;
        CellData[][] result = new CellData[height][];
        for (int r = 0; r < height; r++)
        {
            result[r] = new CellData[width];
            for (int c = 0; c < width; c++) result[r][c] = cells.GetValueOrDefault((r, c), CellData.Plain(string.Empty));
        }
        return new ClipboardTable(result, merges);
    }

    internal static ClipboardTable? FromText(string? text)
    {
        if (string.IsNullOrEmpty(text)) return null;
        List<CellData[]> rows = new();
        List<CellData> row = new();
        StringBuilder value = new();
        bool quoted = false;
        for (int i = 0; i < text.Length; i++)
        {
            char ch = text[i];
            if (ch == '"' && (quoted || value.Length == 0))
            {
                if (quoted && i + 1 < text.Length && text[i + 1] == '"') { value.Append('"'); i++; }
                else quoted = !quoted;
            }
            else if (!quoted && ch is '\t' or '\r' or '\n')
            {
                row.Add(CellData.Plain(value.ToString()));
                value.Clear();
                if (ch != '\t')
                {
                    rows.Add(row.ToArray());
                    row.Clear();
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                }
            }
            else value.Append(ch);
        }
        if (value.Length > 0 || row.Count > 0 || rows.Count == 0 || text[^1] is not ('\r' or '\n'))
        {
            row.Add(CellData.Plain(value.ToString()));
            rows.Add(row.ToArray());
        }
        return new ClipboardTable(rows.ToArray(), Array.Empty<MergeRange>());
    }

    private static CellData ReadCell(HtmlNode cell, List<StyleRule> rules)
    {
        Dictionary<string, string> style = new(StringComparer.OrdinalIgnoreCase);
        HtmlNode? firstText = cell.Descendants().FirstOrDefault(node => node.NodeType == HtmlNodeType.Text && !string.IsNullOrWhiteSpace(WebUtility.HtmlDecode(node.InnerText)));
        IEnumerable<HtmlNode> path = (firstText?.ParentNode ?? cell).AncestorsAndSelf().Reverse();
        foreach (HtmlNode node in path)
        {
            string[] classes = node.GetAttributeValue("class", "").Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            foreach (StyleRule rule in rules)
            {
                if ((rule.Tag.Length == 0 || rule.Tag.Equals(node.Name, StringComparison.OrdinalIgnoreCase)) &&
                    rule.Classes.All(name => classes.Contains(name, StringComparer.Ordinal)))
                    foreach ((string key, string value) in rule.Values) style[key] = value;
            }
            if (node.Name is "b" or "strong" or "th") style["font-weight"] = "bold";
            if (node.Name is "i" or "em") style["font-style"] = "italic";
            if (node.Name == "u") style["text-decoration"] = "underline";
            if (node.Name is "s" or "strike" or "del") style["text-decoration"] = "line-through";
            foreach ((string attribute, string property) in new[] { ("bgcolor", "background-color"), ("color", "color"), ("face", "font-family"), ("align", "text-align"), ("valign", "vertical-align") })
                if (node.Attributes[attribute] is { } attr) style[property] = attr.Value;
            foreach ((string key, string value) in Declarations(node.GetAttributeValue("style", ""))) style[key] = value;
        }

        string Get(string key) => style.GetValueOrDefault(key, "");
        StringBuilder text = new();
        AppendText(cell, text);
        string content = WebUtility.HtmlDecode(text.ToString()).Trim(' ', '\r', '\n').Replace('\u00a0', ' ');
        string? format = style.TryGetValue("mso-number-format", out string? rawFormat) ? Unquote(rawFormat) : null;
        if (cell.Attributes["x:num"] is { Value: { Length: > 0 } numeric } &&
            double.TryParse(numeric, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) && double.IsFinite(number))
            content = number.ToString("R", CultureInfo.InvariantCulture);

        bool underline = Get("text-decoration").Contains("underline", StringComparison.OrdinalIgnoreCase);
        return CellData.Plain(content) with
        {
            Bold = Get("font-weight").Equals("bold", StringComparison.OrdinalIgnoreCase) ||
                int.TryParse(Get("font-weight"), out int weight) && weight >= 600,
            Italic = Get("font-style").Equals("italic", StringComparison.OrdinalIgnoreCase),
            Underline = underline,
            Strikethrough = Get("text-decoration").Contains("line-through", StringComparison.OrdinalIgnoreCase),
            Foreground = ParseColor(Get("color")),
            Background = ParseColor(Get("background-color")) ?? ParseColor(Get("background")),
            FontSize = Length(Get("font-size")),
            FontFamily = Unquote(Get("font-family").Split(',')[0]) is { Length: > 0 } family ? family : null,
            Alignment = Get("text-align").ToLowerInvariant() switch { "center" => TextAlignment.Center, "right" => TextAlignment.Right, _ => TextAlignment.Left },
            NumberFormat = format,
            WrapText = Get("white-space") is "normal" or "pre-wrap" || content.Contains('\n'),
            Appearance = new CellAppearance
            {
                VerticalAlignment = Get("vertical-align").ToLowerInvariant() switch { "top" => VerticalAlignment.Top, "middle" => VerticalAlignment.Center, _ => VerticalAlignment.Bottom },
                DoubleUnderline = underline && Get("text-underline-style").Equals("double", StringComparison.OrdinalIgnoreCase),
                Left = Edge(Get("border-left"), Get("border")), Right = Edge(Get("border-right"), Get("border")),
                Top = Edge(Get("border-top"), Get("border")), Bottom = Edge(Get("border-bottom"), Get("border")),
            },
        };
    }

    private static void AppendText(HtmlNode node, StringBuilder result)
    {
        if (node.Name is "script" or "style") return;
        if (node.Name == "br") { result.Append('\n'); return; }
        if (node.NodeType == HtmlNodeType.Text)
        {
            // Source-code line breaks are HTML whitespace; explicit Excel line breaks use <br>.
            foreach (char ch in node.InnerText) result.Append(ch is '\r' or '\n' or '\t' ? ' ' : ch);
            return;
        }
        foreach (HtmlNode child in node.ChildNodes) AppendText(child, result);
        if (node.Name is "p" or "div" && result.Length > 0 && result[^1] != '\n') result.Append('\n');
    }

    private static Dictionary<string, string> Declarations(string css)
    {
        Dictionary<string, string> values = new(StringComparer.OrdinalIgnoreCase);
        int start = 0;
        char quote = '\0';
        for (int i = 0; i <= css.Length; i++)
        {
            if (i < css.Length)
            {
                char ch = css[i];
                if (ch == '\\') { i++; continue; }
                if (ch == quote) quote = '\0';
                else if (quote == '\0' && ch is '\'' or '"') quote = ch;
                if (ch != ';' || quote != '\0') continue;
            }
            string declaration = css[start..Math.Min(i, css.Length)];
            int colon = declaration.IndexOf(':');
            if (colon > 0) values[declaration[..colon].Trim()] = declaration[(colon + 1)..].Replace("!important", "", StringComparison.OrdinalIgnoreCase).Trim();
            start = i + 1;
        }
        return values;
    }

    private static string Unquote(string value)
    {
        value = value.Trim();
        if (value.Length >= 2 && value[0] is '\'' or '"' && value[^1] == value[0]) value = value[1..^1];
        return Regex.Replace(value, @"\\([0-9a-fA-F]{1,6})\s?|\\(.)", match => match.Groups[1].Success &&
            int.TryParse(match.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code) && Rune.IsValid(code)
                ? char.ConvertFromUtf32(code) : match.Groups[2].Value, RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));
    }

    private static double Length(string value)
    {
        double multiplier = value.EndsWith("pt", StringComparison.OrdinalIgnoreCase) ? 4.0 / 3.0 : 1;
        if (value.EndsWith("pt", StringComparison.OrdinalIgnoreCase) || value.EndsWith("px", StringComparison.OrdinalIgnoreCase)) value = value[..^2];
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double size) && double.IsFinite(size) && size is > 0 and <= 1000 ? size * multiplier : 0;
    }

    private static Color? ParseColor(string text) => Color.TryParse(text.Trim(), out Color color) ? color : null;

    private static CellEdge? Edge(string border, string allBorders)
    {
        if (border.Length == 0) border = allBorders;
        string[] parts = border.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 0 || parts.Contains("none") || parts.Contains("hidden")) return null;
        string style = parts.Contains("double") ? "double" : parts.Contains("dashed") ? "dashed" : parts.Contains("dotted") ? "dotted" :
            Length(parts[0]) >= 2 ? "medium" : "thin";
        return new CellEdge(style, parts.Select(ParseColor).FirstOrDefault(color => color is not null) ?? Colors.Black);
    }

    private sealed record StyleRule(string Tag, string[] Classes, Dictionary<string, string> Values, int Specificity);
}
