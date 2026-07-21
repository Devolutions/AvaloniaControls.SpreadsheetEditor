namespace SpreadsheetEditor.Controls.Formulas;

using System;
using System.Collections.Generic;
using System.Globalization;

// Formats a number with a (subset of) Excel number-format codes — enough to render currency,
// percent, thousands separators, decimal places, and parenthesised negatives faithfully. Falls back
// to "General" for empty/unknown codes. Used so imported/computed numeric cells keep their look
// (e.g. "$"#,##0_);("$"#,##0) → $50,000 / ($5,060)).
public static class NumberFormatter
{
    public static string Format(double value, string? code)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return General(value);
        }

        if (string.IsNullOrWhiteSpace(code) || code.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            return General(value);
        }

        List<string> sections = SplitSections(code!);
        string section;
        if (value < 0 && sections.Count > 1)
        {
            section = sections[1];
        }
        else if (value == 0 && sections.Count > 2)
        {
            section = sections[2];
        }
        else
        {
            section = sections[0];
        }

        try
        {
            return FormatSection(value, section);
        }
        catch
        {
            return General(value);
        }
    }

    private static string FormatSection(double value, string section)
    {
        bool percent = section.Contains('%');
        bool currency = section.Contains('$');
        bool parens = section.Contains('(');
        bool thousands = section.Contains(',');
        int decimals = DecimalPlaces(section);

        double magnitude = Math.Abs(value) * (percent ? 100 : 1);
        string pattern = (thousands ? "#,##0" : "0") + (decimals > 0 ? "." + new string('0', decimals) : string.Empty);
        string number = magnitude.ToString(pattern, CultureInfo.InvariantCulture);

        string body = (currency ? "$" : string.Empty) + number + (percent ? "%" : string.Empty);
        if (value < 0)
        {
            return parens ? $"({body})" : "-" + body;
        }

        return body;
    }

    private static int DecimalPlaces(string section)
    {
        int dot = section.IndexOf('.');
        if (dot < 0)
        {
            return 0;
        }

        int count = 0;
        for (int i = dot + 1; i < section.Length && (section[i] == '0' || section[i] == '#'); i++)
        {
            count++;
        }

        return count;
    }

    // Splits a format code into its positive;negative;zero;text sections, ignoring ';'
    // inside quotes or [brackets].
    private static List<string> SplitSections(string code)
    {
        List<string> sections = new();
        int start = 0;
        bool inQuote = false;
        bool inBracket = false;
        for (int i = 0; i < code.Length; i++)
        {
            char c = code[i];
            if (c == '"')
            {
                inQuote = !inQuote;
            }
            else if (c == '[' && !inQuote)
            {
                inBracket = true;
            }
            else if (c == ']' && !inQuote)
            {
                inBracket = false;
            }
            else if (c == ';' && !inQuote && !inBracket)
            {
                sections.Add(code[start..i]);
                start = i + 1;
            }
        }

        sections.Add(code[start..]);
        return sections;
    }

    // Excel "General": integers without decimals, otherwise up to ~10 trimmed fractional digits.
    private static string General(double value)
    {
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return value.ToString(CultureInfo.InvariantCulture);
        }

        if (value == Math.Floor(value) && Math.Abs(value) < 1e15)
        {
            return ((long)value).ToString(CultureInfo.InvariantCulture);
        }

        return value.ToString("0.##########", CultureInfo.InvariantCulture);
    }
}
