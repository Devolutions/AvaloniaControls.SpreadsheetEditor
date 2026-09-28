namespace SpreadsheetEditor.Controls.ViewModels;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using Avalonia.Media;

using CommunityToolkit.Mvvm.ComponentModel;

using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Models;

// One cell. Holds plain Text + whole-cell formatting flags (used by in-cell plain
// editing, the toolbar, and export), plus an optional list of CellRun for mixed-run
// (per-character) formatting produced by the rich formula bar. When Runs is null the
// cell is uniform; EffectiveRuns always yields the runs to render.
public partial class CellViewModel : ObservableObject
{
    [ObservableProperty]
    private string text = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(FontWeight))]
    private bool bold;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(FontStyle))]
    private bool italic;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(TextDecorations))]
    private bool underline;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(TextDecorations))]
    private bool strikethrough;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Foreground))]
    private Color? foregroundColor;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(Background))]
    private Color? backgroundColor;

    [ObservableProperty]
    private bool isActive;

    // Default cell font size (px). Excel's default 11pt is about 14.67px; we use a slightly tighter default.
    public const double DefaultFontSize = 13;

    [ObservableProperty]
    private double fontSize = DefaultFontSize;

    // Whole-cell font family name (null = inherit the grid default).
    [ObservableProperty]
    private string? fontFamily;

    [ObservableProperty]
    private TextAlignment textAlignment = TextAlignment.Left;

    // The computed result of a formula cell (set by the engine); null for non-formula cells.
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(EffectiveRuns))]
    [NotifyPropertyChangedFor(nameof(DisplayText))]
    private string? computedValue;

    // Excel number-format code for this cell (e.g. "$"#,##0), applied to numeric
    // values for display. Null = show the text as-is.
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(EffectiveRuns))]
    [NotifyPropertyChangedFor(nameof(DisplayText))]
    private string? numberFormat;

    private List<CellRun>? runs;

    private bool applyingRuns;

    // Mixed-run formatting (per-character) from the rich formula bar, or null when uniform.
    public IReadOnlyList<CellRun>? Runs => this.runs;

    // True when Text is a formula (starts with '='). Such cells display their
    // computed result; the raw formula is shown only while editing.
    public bool IsFormula => FormulaEngine.IsFormula(this.Text);

    // What the grid shows: a formula cell's computed result, a numeric cell's value formatted
    // with its number format (e.g. $50,000), otherwise the raw text.
    public string DisplayText
    {
        get
        {
            if (this.IsFormula)
            {
                return this.ComputedValue ?? string.Empty;
            }

            if (this.NumberFormat is { } fmt &&
                double.TryParse(this.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
            {
                return NumberFormatter.Format(n, fmt);
            }

            return this.Text;
        }
    }

    // The runs to render: a single run of the (formatted) display text, unless the cell has
    // explicit mixed runs (rich text), in which case those are used.
    public IReadOnlyList<CellRun> EffectiveRuns =>
        this.IsFormula || this.runs is null
            ? new List<CellRun> { new(this.DisplayText, this.Bold, this.Italic, this.Underline, this.ForegroundColor, this.BackgroundColor, this.Strikethrough) }
            : this.runs;

    // Runs to load into the in-cell editor: the RAW text/formula (so you edit the underlying
    // value, not its formatted display), or the explicit mixed runs for rich text.
    public IReadOnlyList<CellRun> EditRuns =>
        this.runs ?? new List<CellRun> { new(this.Text, this.Bold, this.Italic, this.Underline, this.ForegroundColor, this.BackgroundColor, this.Strikethrough) };

    // Derived, bindable view properties for the active cell's in-cell TextBox + display fallbacks.
    public FontWeight FontWeight => this.Bold ? FontWeight.Bold : FontWeight.Normal;

    public FontStyle FontStyle => this.Italic ? FontStyle.Italic : FontStyle.Normal;

    public TextDecorationCollection? TextDecorations
    {
        get
        {
            if (!this.Underline && !this.Strikethrough)
            {
                return null;
            }

            TextDecorationCollection decorations = new();
            if (this.Underline)
            {
                decorations.Add(new TextDecoration { Location = TextDecorationLocation.Underline });
            }

            if (this.Strikethrough)
            {
                decorations.Add(new TextDecoration { Location = TextDecorationLocation.Strikethrough });
            }

            return decorations;
        }
    }

    public IBrush? Foreground => this.ForegroundColor is { } c ? new SolidColorBrush(c) : null;

    public IBrush? Background => this.BackgroundColor is { } c ? new SolidColorBrush(c) : null;

    public CellViewModel()
    {
    }

    public CellViewModel(string text) => this.text = text;

    // Sets uniform (non-mixed) content from the in-cell editor: whole-cell formatting flags + text,
    // with any explicit runs cleared. Clearing runs lets the cell render via DisplayText, which
    // applies the number format (e.g. $#,##0.00) — so editing a currency cell keeps its $ format,
    // instead of showing the raw typed number. Mixed-format edits still use SetRichRuns.
    public void SetUniformContent(string text, bool bold, bool italic, bool underline, bool strikethrough, Color? foreground)
    {
        this.runs = null;
        this.Bold = bold;
        this.Italic = italic;
        this.Underline = underline;
        this.Strikethrough = strikethrough;
        this.ForegroundColor = foreground;
        this.applyingRuns = true;
        this.Text = text;
        this.applyingRuns = false;
        this.OnPropertyChanged(nameof(this.EffectiveRuns));
    }

    // Sets mixed-run rich content (from the formula bar) and the matching plain text.
    public void SetRichRuns(List<CellRun> runs)
    {
        this.runs = runs.Count == 0 ? null : runs;
        this.applyingRuns = true;
        this.Text = string.Concat(runs.Select(r => r.Text));
        this.applyingRuns = false;
        this.OnPropertyChanged(nameof(this.EffectiveRuns));
    }

    // ----- Whole-cell formatting (applied when the cell is selected, not being edited) -----

    // Whole-cell formatting state (true only if uniform across runs).
    public bool IsAllBold => this.IsFormula ? this.Bold : this.runs?.All(r => r.Bold) ?? this.Bold;

    public bool IsAllItalic => this.IsFormula ? this.Italic : this.runs?.All(r => r.Italic) ?? this.Italic;

    public bool IsAllUnderline => this.IsFormula ? this.Underline : this.runs?.All(r => r.Underline) ?? this.Underline;

    public bool IsAllStrikethrough => this.IsFormula ? this.Strikethrough : this.runs?.All(r => r.Strikethrough) ?? this.Strikethrough;

    public void SetWholeBold(bool on) => this.ApplyToWholeCell(r => r with { Bold = on }, () => this.Bold = on);

    public void SetWholeItalic(bool on) => this.ApplyToWholeCell(r => r with { Italic = on }, () => this.Italic = on);

    public void SetWholeUnderline(bool on) => this.ApplyToWholeCell(r => r with { Underline = on }, () => this.Underline = on);

    public void SetWholeStrikethrough(bool on) => this.ApplyToWholeCell(r => r with { Strikethrough = on }, () => this.Strikethrough = on);

    public void SetWholeForeground(Color color) =>
        this.ApplyToWholeCell(r => r with { Foreground = color }, () => this.ForegroundColor = color);

    public void ToggleWholeBold() => this.SetWholeBold(!this.IsAllBold);

    public void ToggleWholeItalic() => this.SetWholeItalic(!this.IsAllItalic);

    public void ToggleWholeUnderline() => this.SetWholeUnderline(!this.IsAllUnderline);

    public void ToggleWholeStrikethrough() => this.SetWholeStrikethrough(!this.IsAllStrikethrough);

    // Applies a transform to every run (preserving other per-run formatting), or just sets
    // the whole-cell flag for formula cells (whose display comes from the computed value + flags).
    private void ApplyToWholeCell(Func<CellRun, CellRun> transform, Action setFlag)
    {
        setFlag();
        if (this.IsFormula)
        {
            // formula display uses the flags directly (EffectiveRuns ignores runs)
            return;
        }

        this.SetRichRuns(this.EffectiveRuns.Select(transform).ToList());
    }

    // In-cell plain typing flattens any mixed-run formatting; whole-cell flag changes re-render.
    partial void OnTextChanged(string value)
    {
        if (!this.applyingRuns)
        {
            this.runs = null;
        }

        this.OnPropertyChanged(nameof(IsFormula));
        this.OnPropertyChanged(nameof(DisplayText));
        this.OnPropertyChanged(nameof(EffectiveRuns));
    }

    partial void OnBoldChanged(bool value) => this.NotifyEffectiveRuns();

    partial void OnItalicChanged(bool value) => this.NotifyEffectiveRuns();

    partial void OnUnderlineChanged(bool value) => this.NotifyEffectiveRuns();

    partial void OnStrikethroughChanged(bool value) => this.NotifyEffectiveRuns();

    partial void OnForegroundColorChanged(Color? value) => this.NotifyEffectiveRuns();

    partial void OnBackgroundColorChanged(Color? value) => this.NotifyEffectiveRuns();

    private void NotifyEffectiveRuns()
    {
        if (this.runs is null)
        {
            this.OnPropertyChanged(nameof(this.EffectiveRuns));
        }
    }

    public CellData ToData() =>
        new(
            this.Text,
            this.Bold,
            this.Italic,
            this.Underline,
            this.ForegroundColor,
            this.BackgroundColor,
            this.FontSize,
            this.TextAlignment,
            this.NumberFormat,
            this.FontFamily,
            this.Strikethrough,
            this.IsFormula ? this.ComputedValue : null);

    public void ApplyData(CellData data)
    {
        this.runs = null;
        this.Text = data.Text;
        this.Bold = data.Bold;
        this.Italic = data.Italic;
        this.Underline = data.Underline;
        this.Strikethrough = data.Strikethrough;
        this.ForegroundColor = data.Foreground;
        this.BackgroundColor = data.Background;
        this.FontSize = data.FontSize > 0 ? data.FontSize : DefaultFontSize;
        this.TextAlignment = data.Alignment;
        this.NumberFormat = data.NumberFormat;
        this.FontFamily = data.FontFamily;
        this.ComputedValue = FormulaEngine.IsFormula(data.Text) ? data.FormulaResult : null;
        this.OnPropertyChanged(nameof(this.EffectiveRuns));
    }
}