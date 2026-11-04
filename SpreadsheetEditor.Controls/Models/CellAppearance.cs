namespace SpreadsheetEditor.Controls.Models;

using System.Collections.Generic;
using Avalonia.Layout;
using Avalonia.Media;

// The less common Excel presentation options travel with the cell through copy, undo and save.
public sealed record CellAppearance
{
    public VerticalAlignment VerticalAlignment { get; init; } = VerticalAlignment.Bottom;
    public int Rotation { get; init; }
    public int Indent { get; init; }
    public bool ShrinkToFit { get; init; }
    public bool DoubleUnderline { get; init; }
    public string? Script { get; init; }
    public bool Locked { get; init; } = true;
    public CellEdge? Left { get; init; }
    public CellEdge? Right { get; init; }
    public CellEdge? Top { get; init; }
    public CellEdge? Bottom { get; init; }
    public CellFill? Fill { get; init; }
}

public sealed record CellEdge(string Style, Color Color);
public sealed record CellGradientStop(double Position, Color Color);
public sealed record CellFill
{
    public string? Pattern { get; init; }
    public Color? Foreground { get; init; }
    public Color? Background { get; init; }
    public double Degree { get; init; }
    public IReadOnlyList<CellGradientStop> Stops { get; init; } = System.Array.Empty<CellGradientStop>();
}

public sealed record SheetOutlineGroup(int Start, int Count, int Level, bool Collapsed, bool SummaryAfter);
