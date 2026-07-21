namespace SpreadsheetEditor.Controls.Models;

using Avalonia.Media;

// A contiguous run of text within a cell sharing the same formatting. A cell's rich content is a
// sequence of these (mixed-run / per-character formatting). A uniform cell collapses to one run.
public readonly record struct CellRun(
    string Text,
    bool Bold,
    bool Italic,
    bool Underline,
    Color? Foreground,
    Color? Background,
    bool Strikethrough = false);
