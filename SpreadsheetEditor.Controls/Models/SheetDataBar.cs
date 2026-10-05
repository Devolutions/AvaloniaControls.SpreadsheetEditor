namespace SpreadsheetEditor.Controls.Models;

using Avalonia.Media;

// A conditional-formatting data bar over a cell range: each cell shows a horizontal bar whose width
// is proportional to its value between Min and Max. When ShowValue is false, only the bar is shown
// (the number is hidden), like Excel.
public sealed class SheetDataBar
{
    public required CellRange Range { get; init; }

    public required double Min { get; init; }

    public required double Max { get; init; }

    public required Color Color { get; init; }

    public bool ShowValue { get; init; } = true;

    // Standard Excel data bars use a left-to-right gradient unless a solid fill is requested.
    public bool Gradient { get; init; } = true;
}
