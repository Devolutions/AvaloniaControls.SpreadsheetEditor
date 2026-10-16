namespace SpreadsheetEditor.Controls.Models;

using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Formulas;

// A transport record for a cell's value + whole-cell formatting. Used to move data between the
// control and host-app serializers (TSV plain text, or XLSX with styles) without exposing the
// view models. FormulaResult can carry an imported workbook's saved display value, avoiding a full
// recalculation merely to paint the sheet. Whole-cell (not mixed-run) formatting is the v1 fidelity level.
public readonly record struct CellData(
    string Text,
    bool Bold,
    bool Italic,
    bool Underline,
    Color? Foreground,
    Color? Background,
    double FontSize = 0,
    TextAlignment Alignment = TextAlignment.Left,
    string? NumberFormat = null,
    string? FontFamily = null,
    bool Strikethrough = false,
    string? FormulaResult = null,
    string? Comment = null,
    string? Hyperlink = null,
    string? CommentAuthor = null,
    double? FormulaNumericValue = null)
{
    public bool WrapText { get; init; }

    // The unformatted cached result keeps numbers, text, booleans and errors distinct on save.
    public CellValue? FormulaValue { get; init; }

    public static CellData Plain(string text) => new(text, false, false, false, null, null);
}
