namespace SpreadsheetEditor.Controls;

using System;

using Avalonia;
using Avalonia.Input;

using global::SpreadsheetEditor.Controls.ViewModels;

public partial class SpreadsheetGrid
{
    private sealed record HeaderResizeDrag(
        SpreadsheetViewModel Sheet, IPointer Pointer, bool IsColumn, int Index,
        double StartCoordinate, double StartSize, double? OriginalRowHeight, Cursor? OriginalCursor);

    private HeaderResizeDrag? headerResize;

    private void InitializeHeaderResize()
    {
        this.PointerMoved += (_, e) =>
        {
            if (this.headerResize is not { } drag || !ReferenceEquals(e.Pointer.Captured, this)) return;
            this.UpdateHeaderResize(drag, e.GetPosition(this));
            e.Handled = true;
        };
        this.PointerReleased += (_, e) =>
        {
            if (this.headerResize is not { } drag || !ReferenceEquals(e.Pointer.Captured, this) ||
                e.InitialPressMouseButton != MouseButton.Left) return;
            // The last movement may arrive with the release rather than a separate move event.
            this.UpdateHeaderResize(drag, e.GetPosition(this));
            this.FinishHeaderResize();
            e.Handled = true;
        };
        this.PointerCaptureLost += (_, _) =>
        {
            if (this.headerResize is { } drag && !ReferenceEquals(drag.Pointer.Captured, this))
            {
                this.FinishHeaderResize();
            }
        };
        this.DetachedFromVisualTree += (_, _) => this.FinishHeaderResize(cancel: true, rebuild: false);
    }

    private void BeginHeaderResize(PointerPressedEventArgs e, int index, bool isColumn)
    {
        if (this.ReadOnly || this.sheet is null) return;
        this.CancelFillDrag();
        this.SuspendEditor();
        this.Focus();
        Point position = e.GetPosition(this);
        this.headerResize = new HeaderResizeDrag(this.sheet, e.Pointer, isColumn, index,
            isColumn ? position.X : position.Y,
            isColumn ? this.columnWidths[index] : this.rowHeights[index],
            isColumn ? null : this.sheet.Rows[index].Height, this.Cursor);
        this.Cursor = new Cursor(isColumn ? StandardCursorType.SizeWestEast : StandardCursorType.SizeNorthSouth);
        // Viewport rendering replaces the header controls on every size update. Capture the
        // persistent grid so a repaint cannot terminate the drag or leave resizing stuck on.
        e.Pointer.Capture(this);
        e.Handled = true;
    }

    private void UpdateHeaderResize(HeaderResizeDrag drag, Point position)
    {
        if (!ReferenceEquals(this.sheet, drag.Sheet)) return;
        double coordinate = drag.IsColumn ? position.X : position.Y;
        double size = Math.Max((drag.IsColumn ? MinColumnWidth : MinRowHeight) * this.zoom,
            drag.StartSize + coordinate - drag.StartCoordinate);
        double current = drag.IsColumn ? this.columnWidths[drag.Index] : this.rowHeights[drag.Index];
        if (Math.Abs(size - current) < 0.001) return;
        if (drag.IsColumn) this.SetColumnWidth(drag.Index, size);
        else this.SetRowHeight(drag.Index, size);
    }

    private void FinishHeaderResize(bool cancel = false, bool rebuild = true)
    {
        if (this.headerResize is not { } drag) return;
        this.headerResize = null;
        this.Cursor = drag.OriginalCursor;
        bool sameSheet = ReferenceEquals(this.sheet, drag.Sheet);
        bool changed = sameSheet && Math.Abs((drag.IsColumn
            ? this.columnWidths[drag.Index] : this.rowHeights[drag.Index]) - drag.StartSize) >= 0.001;
        if (cancel && sameSheet)
        {
            if (drag.IsColumn) this.SetColumnWidth(drag.Index, drag.StartSize);
            else
            {
                this.SetRowHeight(drag.Index, drag.StartSize);
                drag.Sheet.Rows[drag.Index].Height = drag.OriginalRowHeight;
            }
        }
        if (ReferenceEquals(drag.Pointer.Captured, this)) drag.Pointer.Capture(null);
        if (sameSheet && rebuild) this.Rebuild();
        if (sameSheet && changed && !cancel) drag.Sheet.NotifyMetadataChanged();
    }
}
