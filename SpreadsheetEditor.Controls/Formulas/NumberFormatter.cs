namespace SpreadsheetEditor.Controls.Formulas;

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Avalonia.Media;
using ExcelNumberFormat;

// ClosedXML's ECMA-376 formatter also handles dates, fractions, literals, scientific notation,
// locale currency symbols and the positive/negative/zero format sections.
public static class NumberFormatter
{
    private static readonly ConcurrentDictionary<string, NumberFormat> Formats = new(StringComparer.Ordinal);

    public static string Format(double value, string? code)
    {
        if (string.IsNullOrWhiteSpace(code) || code.Equals("General", StringComparison.OrdinalIgnoreCase) || !double.IsFinite(value))
            return value.ToString("0.##########", CultureInfo.InvariantCulture);
        NumberFormat format = Formats.GetOrAdd(code, static c => new NumberFormat(c));
        return format.IsValid ? format.Format(value, CultureInfo.InvariantCulture) : value.ToString(CultureInfo.InvariantCulture);
    }

    public static Color? Foreground(double value, string? code)
    {
        if (string.IsNullOrEmpty(code) || !code.Contains('[')) return null;
        List<string> sections = SplitSections(code);
        string section = sections[value < 0 && sections.Count > 1 ? 1 : value == 0 && sections.Count > 2 ? 2 : 0];
        foreach (string candidate in sections)
        {
            Match condition = Regex.Match(candidate, @"\[(<=|>=|<>|=|<|>)(-?[0-9.]+)\]");
            if (!condition.Success) continue;
            if (!double.TryParse(condition.Groups[2].Value, CultureInfo.InvariantCulture, out double threshold)) continue;
            bool matches = condition.Groups[1].Value switch
            {
                "<=" => value <= threshold, ">=" => value >= threshold, "<>" => value != threshold,
                "=" => value == threshold, "<" => value < threshold, _ => value > threshold,
            };
            if (matches) { section = candidate; break; }
        }
        Match color = Regex.Match(section, @"\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]", RegexOptions.IgnoreCase);
        return color.Success ? Color.Parse(color.Groups[1].Value) : null;
    }

    private static List<string> SplitSections(string code)
    {
        List<string> result = new();
        bool quoted = false;
        int start = 0;
        for (int i = 0; i < code.Length; i++)
        {
            if (code[i] == '\\') { i++; continue; }
            if (code[i] == '"') quoted = !quoted;
            if (code[i] == ';' && !quoted) { result.Add(code[start..i]); start = i + 1; }
        }
        result.Add(code[start..]);
        return result;
    }

    public static string? AccountingPrefix(string? code)
    {
        if (code is null || !code.Contains('*')) return null;
        Match localeCurrency = Regex.Match(code, @"\[\$([^\]-]+)(?:-[^\]]+)?\]");
        if (localeCurrency.Success) return localeCurrency.Groups[1].Value;
        Match symbol = Regex.Match(code, @"\p{Sc}");
        return symbol.Success ? symbol.Value : null;
    }
}
