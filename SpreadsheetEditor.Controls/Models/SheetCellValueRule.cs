namespace SpreadsheetEditor.Controls.Models;

using Avalonia.Media;

// A numeric cell-is conditional font rule. Kept as a rule rather than painting imported cells so
// the text changes immediately when the user edits a value or a dependent formula recalculates.
public enum CellValueComparison
{
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
}

public sealed class SheetCellValueRule
{
    public required CellRange Range { get; init; }

    public required CellValueComparison Comparison { get; init; }

    public required double Threshold { get; init; }

    public Color? Foreground { get; init; }

    public bool? Bold { get; init; }

    public int Priority { get; init; }

    public bool Matches(double value) => this.Comparison switch
    {
        CellValueComparison.GreaterThan => value > this.Threshold,
        CellValueComparison.GreaterThanOrEqual => value >= this.Threshold,
        CellValueComparison.LessThan => value < this.Threshold,
        CellValueComparison.LessThanOrEqual => value <= this.Threshold,
        CellValueComparison.Equal => value == this.Threshold,
        CellValueComparison.NotEqual => value != this.Threshold,
        _ => false,
    };
}
