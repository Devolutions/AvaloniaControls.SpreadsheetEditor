namespace SpreadsheetEditor.Controls.Models;

// A merged cell region: the top-left cell (zero-based) and how many rows/columns it spans.
public readonly record struct MergeRange(int Row, int Column, int RowSpan, int ColumnSpan);
