namespace SpreadsheetEditor.Controls.Formulas;

using System;
using System.Globalization;

// Excel-style error literals produced by the formula engine.
public static class FormulaErrors
{
    public const string Div0 = "#DIV/0!";
    public const string Value = "#VALUE!";
    public const string Name = "#NAME?";
    public const string Ref = "#REF!";
    public const string Num = "#NUM!";
    public const string NA = "#N/A";
    public const string Circular = "#CIRC!";
}

public enum CellValueKind
{
    Empty,
    Number,
    Text,
    Boolean,
    Error,
}

// A computed cell value: a number, text, boolean, error, or empty. Carries the Excel coercion rules
// (text→number where parseable, bool→1/0, empty→0, errors propagate) and "General"-format display.
public readonly struct CellValue
{
    private CellValue(CellValueKind kind, double number, string text, bool boolean)
    {
        this.Kind = kind;
        this.Number = number;
        this.Text = text;
        this.Boolean = boolean;
    }

    public static readonly CellValue Empty = new(CellValueKind.Empty, 0, string.Empty, false);

    public CellValueKind Kind { get; }

    public double Number { get; }

    // value text, or the error literal when Kind == Error
    public string Text { get; }

    public bool Boolean { get; }

    public bool IsError => this.Kind == CellValueKind.Error;

    public static CellValue FromNumber(double n) => new(CellValueKind.Number, n, string.Empty, false);

    public static CellValue FromText(string s) => new(CellValueKind.Text, 0, s ?? string.Empty, false);

    public static CellValue FromBool(bool b) => new(CellValueKind.Boolean, 0, string.Empty, b);

    public static CellValue FromError(string e) => new(CellValueKind.Error, 0, e, false);

    // Coerces to a number using Excel rules; returns false (with a #VALUE!-worthy result)
    // for non-numeric text. Errors return false too (callers should propagate the error).
    public bool TryAsNumber(out double number)
    {
        switch (this.Kind)
        {
            case CellValueKind.Number:
                number = this.Number;
                return true;
            case CellValueKind.Boolean:
                number = this.Boolean ? 1 : 0;
                return true;
            case CellValueKind.Empty:
                number = 0;
                return true;
            case CellValueKind.Text:
                return double.TryParse(this.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out number);
            default:
                number = 0;
                return false;
        }
    }

    // Coerces to text (numbers use General format, booleans TRUE/FALSE, empty = "").
    public string AsText() => this.Kind switch
    {
        CellValueKind.Text => this.Text,
        CellValueKind.Number => FormatNumber(this.Number),
        CellValueKind.Boolean => this.Boolean ? "TRUE" : "FALSE",
        CellValueKind.Error => this.Text,
        _ => string.Empty,
    };

    // The string shown in a cell for this value.
    public string Display() => this.AsText();

    // Excel "General" number formatting: integers without decimals, otherwise up to ~10
    // fractional digits with trailing zeros trimmed. Invariant culture (engine-internal).
    private static string FormatNumber(double d)
    {
        if (double.IsNaN(d))
        {
            return FormulaErrors.Num;
        }

        if (double.IsInfinity(d))
        {
            return FormulaErrors.Div0;
        }

        if (d == Math.Floor(d) && Math.Abs(d) < 1e15)
        {
            return ((long)d).ToString(CultureInfo.InvariantCulture);
        }

        return d.ToString("0.##########", CultureInfo.InvariantCulture);
    }
}
