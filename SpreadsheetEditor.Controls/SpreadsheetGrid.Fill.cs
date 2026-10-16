namespace SpreadsheetEditor.Controls;

using System;
using System.Linq;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;

using global::SpreadsheetEditor.Controls.Models;

public partial class SpreadsheetGrid
{
    private readonly Border fillHandle = new()
    {
        Name = "FillHandle",
        Width = 12,
        Height = 12,
        Background = Brushes.Transparent,
        Cursor = new Cursor(StandardCursorType.Cross),
        IsVisible = false,
        ZIndex = 10,
        Child = new Border
        {
            Width = 7,
            Height = 7,
            Background = SelectionBorder,
            BorderBrush = SheetSurface,
            BorderThickness = new Thickness(1),
            HorizontalAlignment = HorizontalAlignment.Center,
            VerticalAlignment = VerticalAlignment.Center,
        },
    };

    private readonly Border fillPreview = new()
    {
        Name = "FillPreview",
        BorderBrush = SelectionBorder,
        BorderThickness = new Thickness(2),
        Opacity = 0.7,
        IsHitTestVisible = false,
        IsVisible = false,
        ZIndex = 9,
    };

    private CellRange? fillDragSource;
    private CellRange fillDragTarget;
    private IPointer? fillPointer;
    private Point fillPointerPosition;
    private Point fillStartPosition;
    private bool fillMoved;
    private DispatcherTimer? fillScrollTimer;

    private void InitializeFillHandle()
    {
        this.cells.Children.Add(this.fillPreview);
        this.cells.Children.Add(this.fillHandle);
        this.fillHandle.PointerPressed += (_, e) =>
        {
            if (this.ReadOnly || this.sheet is null || !this.editorSuspended ||
                !e.GetCurrentPoint(this.fillHandle).Properties.IsLeftButtonPressed)
            {
                return;
            }
            this.fillDragSource = this.SelectionRange;
            this.fillDragTarget = this.SelectionRange;
            this.fillStartPosition = this.fillPointerPosition = e.GetPosition(this.cellScroller);
            this.fillMoved = false;
            this.fillPointer = e.Pointer;
            // Capture the persistent grid, not a cell/handle that viewport rendering can detach.
            e.Pointer.Capture(this);
            this.Focus();
            e.Handled = true;
        };
        this.PointerMoved += (_, e) =>
        {
            if (this.fillDragSource is null || !ReferenceEquals(e.Pointer.Captured, this)) return;
            this.fillPointerPosition = e.GetPosition(this.cellScroller);
            if (!this.fillMoved && Math.Max(Math.Abs(this.fillPointerPosition.X - this.fillStartPosition.X),
                    Math.Abs(this.fillPointerPosition.Y - this.fillStartPosition.Y)) < 4) return;
            this.fillMoved = true;
            this.UpdateFillDragTarget();
            this.fillScrollTimer ??= new DispatcherTimer(TimeSpan.FromMilliseconds(30), DispatcherPriority.Background,
                (_, _) => this.ScrollDuringFill());
            this.fillScrollTimer.Start();
            e.Handled = true;
        };
        this.PointerReleased += (_, e) =>
        {
            if (this.fillDragSource is not { } source || !ReferenceEquals(e.Pointer.Captured, this)) return;
            if (this.fillMoved)
            {
                this.fillPointerPosition = e.GetPosition(this.cellScroller);
                this.UpdateFillDragTarget();
            }
            CellRange target = this.fillDragTarget;
            bool apply = this.fillMoved && !this.ReadOnly;
            this.CancelFillDrag();
            if (apply && this.sheet is { } sheet)
            {
                this.ApplyCalculatedCells(() =>
                {
                    if (!sheet.FillCells(source, target)) return;
                    this.rangeAnchor = new CellAddress(target.Row, target.Column);
                    this.rangeFocus = new CellAddress(target.Row + target.RowSpan - 1, target.Column + target.ColumnSpan - 1);
                    this.Rebuild();
                });
            }
            this.Focus();
            e.Handled = true;
        };
        this.PointerCaptureLost += (_, _) =>
        {
            if (this.fillPointer is not null && !ReferenceEquals(this.fillPointer.Captured, this)) this.CancelFillDrag();
        };
        this.DetachedFromVisualTree += (_, _) => this.CancelFillDrag();
    }

    private void UpdateFillHandle(CellRange range, bool editing)
    {
        bool merged = this.sheet?.Merges.Any(merge => range.Row < merge.Row + merge.RowSpan &&
            range.Row + range.RowSpan > merge.Row && range.Column < merge.Column + merge.ColumnSpan &&
            range.Column + range.ColumnSpan > merge.Column) == true;
        this.fillHandle.IsVisible = !this.ReadOnly && !editing && this.selectedChart is null && !merged;
        Canvas.SetLeft(this.fillHandle, this.OffsetToColumn(range.Column + range.ColumnSpan) - this.fillHandle.Width / 2);
        Canvas.SetTop(this.fillHandle, this.OffsetToRow(range.Row + range.RowSpan) - this.fillHandle.Height / 2);
        if (this.fillDragSource is not null)
        {
            this.ShowFillPreview();
        }
    }

    private void UpdateFillDragTarget()
    {
        if (this.fillDragSource is not { } source || this.sheet is null) return;
        Vector offset = this.cellScroller.Offset;
        Point position = new(offset.X + Math.Clamp(this.fillPointerPosition.X, 0, Math.Max(0, this.cellScroller.Viewport.Width - 1)),
            offset.Y + Math.Clamp(this.fillPointerPosition.Y, 0, Math.Max(0, this.cellScroller.Viewport.Height - 1)));
        int row = this.RowAtY(position.Y);
        int column = this.ColAtX(position.X);
        int bottom = source.Row + source.RowSpan - 1;
        int right = source.Column + source.ColumnSpan - 1;
        double horizontal = Math.Max(this.OffsetToColumn(source.Column) - position.X,
            position.X - this.OffsetToColumn(right + 1));
        double vertical = Math.Max(this.OffsetToRow(source.Row) - position.Y,
            position.Y - this.OffsetToRow(bottom + 1));
        // Fill along one axis, including up/left; a diagonal gesture follows its larger extension.
        this.fillDragTarget = horizontal > vertical && horizontal > 0
            ? new CellRange(source.Row, Math.Min(column, source.Column), source.RowSpan,
                Math.Max(column, right) - Math.Min(column, source.Column) + 1)
            : vertical > 0
                ? new CellRange(Math.Min(row, source.Row), source.Column,
                    Math.Max(row, bottom) - Math.Min(row, source.Row) + 1, source.ColumnSpan)
                : source;
        this.ShowFillPreview();
    }

    private void ShowFillPreview()
    {
        CellRange target = this.fillDragTarget;
        this.PlaceSheetControl(this.fillPreview, target.Row, target.Column, target.RowSpan, target.ColumnSpan);
        this.fillPreview.IsVisible = this.fillMoved && target != this.fillDragSource;
    }

    private void ScrollDuringFill()
    {
        if (this.fillDragSource is null || this.ReadOnly)
        {
            this.CancelFillDrag();
            return;
        }
        static double Step(double position, double extent) => position < 20 ? -Math.Clamp(20 - position, 4, 40) :
            position > extent - 20 ? Math.Clamp(position - extent + 20, 4, 40) : 0;
        double dx = Step(this.fillPointerPosition.X, this.cellScroller.Viewport.Width);
        double dy = Step(this.fillPointerPosition.Y, this.cellScroller.Viewport.Height);
        if (dx == 0 && dy == 0) return;
        this.cellScroller.Offset = new Vector(
            Math.Clamp(this.cellScroller.Offset.X + dx, 0, Math.Max(0, this.cellScroller.Extent.Width - this.cellScroller.Viewport.Width)),
            Math.Clamp(this.cellScroller.Offset.Y + dy, 0, Math.Max(0, this.cellScroller.Extent.Height - this.cellScroller.Viewport.Height)));
        this.AdjustRowsForScroll();
        this.AdjustColumnsForScroll();
        this.UpdateFillDragTarget();
    }

    private void CancelFillDrag()
    {
        this.fillScrollTimer?.Stop();
        this.fillDragSource = null;
        this.fillPreview.IsVisible = false;
        IPointer? pointer = this.fillPointer;
        this.fillPointer = null;
        if (ReferenceEquals(pointer?.Captured, this)) pointer.Capture(null);
    }
}
