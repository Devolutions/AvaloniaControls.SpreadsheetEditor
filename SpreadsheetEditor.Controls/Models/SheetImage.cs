namespace SpreadsheetEditor.Controls.Models;

// A picture embedded in a sheet. Positioned by an anchor cell (0-based) plus a pixel offset within
// that cell's top-left, sized in device pixels. Free-floating pictures are normalized to an A1
// anchor (row/column 0) with the offset carrying the absolute position. Data is the raw encoded
// image bytes (PNG/JPEG/etc.).
public sealed class SheetImage
{
    public required int Row { get; init; }

    public required int Column { get; init; }

    public int OffsetX { get; init; }

    public int OffsetY { get; init; }

    public required int Width { get; init; }

    public required int Height { get; init; }

    public required byte[] Data { get; init; }
}
