namespace SpreadsheetEditor.Controls;

using System;
using System.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;

public partial class SpreadsheetEditor
{
    private string? tabStripActiveName;
    private int tabStripGeneration;
    private bool revealActiveSheetTab;

    private void InitializeSheetTabScrolling()
    {
        this.ScrollSheetTabsLeftButton.Click += (_, _) => this.ScrollSheetTabs(forward: false);
        this.ScrollSheetTabsRightButton.Click += (_, _) => this.ScrollSheetTabs(forward: true);
        this.SheetTabScroller.ScrollChanged += (_, _) => this.UpdateSheetTabArrows();
        this.SheetTabScroller.SizeChanged += (_, _) =>
        {
            if (this.revealActiveSheetTab) this.ScheduleSheetTabLayout();
        };
        this.SheetTabScroller.AddHandler(PointerWheelChangedEvent, (_, e) =>
        {
            if (this.SheetTabScroller.Extent.Width <= this.SheetTabScroller.Viewport.Width) return;
            double delta = e.Delta.X != 0 ? e.Delta.X : e.Delta.Y;
            this.SetSheetTabOffset(this.SheetTabScroller.Offset.X - delta * 80);
            e.Handled = true;
        }, RoutingStrategies.Tunnel);
    }

    private void ResetSheetTabScroll()
    {
        this.tabStripActiveName = null;
        this.SheetTabScroller.Offset = default;
    }

    private void UpdateSheetTabArrows()
    {
        double maximum = Math.Max(0, this.SheetTabScroller.Extent.Width - this.SheetTabScroller.Viewport.Width);
        this.ScrollSheetTabsLeftButton.IsEnabled = this.SheetTabScroller.Offset.X > 0.5;
        this.ScrollSheetTabsRightButton.IsEnabled = this.SheetTabScroller.Offset.X < maximum - 0.5;
    }

    private void SetSheetTabOffset(double offset)
    {
        double maximum = Math.Max(0, this.SheetTabScroller.Extent.Width - this.SheetTabScroller.Viewport.Width);
        this.SheetTabScroller.Offset = new Vector(Math.Clamp(offset, 0, maximum), 0);
        this.UpdateSheetTabArrows();
    }

    private void ScrollSheetTabs(bool forward)
    {
        double offset = this.SheetTabScroller.Offset.X;
        double[] starts = this.SheetTabs.Children.Select(tab => tab.Bounds.Left).ToArray();
        double target = forward
            ? starts.FirstOrDefault(x => x > offset + 0.5, this.SheetTabScroller.Extent.Width)
            : starts.LastOrDefault(x => x < offset - 0.5, 0);
        this.SetSheetTabOffset(target);
    }

    private void ScheduleSheetTabLayout()
    {
        string activeName = this.pendingSheetActivationIndex is { } index
            ? this.sheets.FirstOrDefault(sheet => this.stagedSheetIndexes.TryGetValue(sheet, out int stagedIndex) && stagedIndex == index)?.Name ?? this.vm.Name
            : this.vm.Name;
        this.revealActiveSheetTab |= this.tabStripActiveName != activeName;
        this.tabStripActiveName = activeName;
        int generation = ++this.tabStripGeneration;
        Dispatcher.UIThread.Post(() =>
        {
            if (generation != this.tabStripGeneration) return;
            if (this.revealActiveSheetTab && this.SheetTabScroller.Viewport.Width > 0 &&
                this.SheetTabs.Children.OfType<Button>().FirstOrDefault(tab => Equals(tab.Tag, activeName)) is { Bounds.Width: > 0 } active)
            {
                this.revealActiveSheetTab = false;
                double offset = this.SheetTabScroller.Offset.X;
                if (active.Bounds.Left < offset) this.SetSheetTabOffset(active.Bounds.Left);
                else if (active.Bounds.Right > offset + this.SheetTabScroller.Viewport.Width)
                    this.SetSheetTabOffset(active.Bounds.Width > this.SheetTabScroller.Viewport.Width
                        ? active.Bounds.Left : active.Bounds.Right - this.SheetTabScroller.Viewport.Width);
            }
            this.UpdateSheetTabArrows();
        }, DispatcherPriority.Loaded);
    }
}
