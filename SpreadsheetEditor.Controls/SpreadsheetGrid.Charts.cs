namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Threading;

using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.ViewModels;

public partial class SpreadsheetGrid
{
    private sealed record ChartData(List<string> Categories, List<ChartSeriesData> Series);
    private readonly Dictionary<SheetChart, ChartData> renderedChartData = new();
    private CancellationTokenSource? chartRefreshCancellation;
    private bool chartRefreshDeferred;

    private static bool SameChartData(ChartData first, ChartData second) =>
        first.Categories.SequenceEqual(second.Categories) && first.Series.Count == second.Series.Count &&
        first.Series.Zip(second.Series).All(pair => pair.First.Title == pair.Second.Title &&
            pair.First.Values.SequenceEqual(pair.Second.Values));

    private void RefreshCharts() => this.ScheduleChartRefresh();

    private void ScheduleChartRefresh()
    {
        this.CancelChartRefresh();
        if (this.chartDrag is not null)
        {
            this.chartRefreshDeferred = true;
            return;
        }
        if (this.sheet is null) return;
        if (this.sheet.Charts.Count == 0)
        {
            this.chartOverlay.Children.Clear();
            this.chartCards.Clear();
            this.renderedChartData.Clear();
            this.selectedChart = null;
            return;
        }
        CancellationTokenSource cancellation = new();
        this.chartRefreshCancellation = cancellation;
        _ = this.RefreshChartsAsync(this.sheet, cancellation);
    }

    private void CancelChartRefresh()
    {
        // Source ownership stays on the UI thread. The worker only retains its token; it must not
        // post a final UI callback after the grid/window has already been detached or replaced.
        CancellationTokenSource? cancellation = this.chartRefreshCancellation;
        this.chartRefreshCancellation = null;
        cancellation?.Cancel();
        cancellation?.Dispose();
    }

    private async Task RefreshChartsAsync(SpreadsheetViewModel model, CancellationTokenSource cancellation)
    {
        CancellationToken token = cancellation.Token;
        try
        {
            // Coalesce repeated change/layout notifications and let input/painting run first.
            await Task.Delay(1, token).ConfigureAwait(false);
            SheetChart[] charts = Array.Empty<SheetChart>();
            int rows = 0, columns = 0;
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                token.ThrowIfCancellationRequested();
                charts = model.Charts.ToArray();
                rows = model.RowCount;
                columns = model.ColumnCount;
                foreach (var entry in this.chartCards.Where(entry => !charts.Contains(entry.Card.Tag)).ToArray())
                {
                    this.chartOverlay.Children.Remove(entry.Card);
                    this.chartCards.Remove(entry);
                    if (entry.Card.Tag is SheetChart removed) this.renderedChartData.Remove(removed);
                }
                if (this.selectedChart is not null && !charts.Contains(this.selectedChart)) this.selectedChart = null;
            }, DispatcherPriority.Background);

            foreach (SheetChart chart in charts)
            {
                ChartData? previous = null;
                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    token.ThrowIfCancellationRequested();
                    this.renderedChartData.TryGetValue(chart, out previous);
                }, DispatcherPriority.Background);
                Dictionary<CellAddress, string> values = new();
                IEnumerable<CellRange> ranges = chart.Series.SelectMany(series => series.Ranges);
                if (chart.Categories is { } categories) ranges = ranges.Append(categories);
                using IEnumerator<CellAddress> addresses = ranges
                    .SelectMany(range => EnumerateRange(rows, columns, range)).GetEnumerator();
                bool more = true;
                while (more)
                {
                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        token.ThrowIfCancellationRequested();
                        Stopwatch budget = Stopwatch.StartNew();
                        while ((more = addresses.MoveNext()))
                        {
                            CellAddress address = addresses.Current;
                            if (!values.ContainsKey(address)) values[address] = model.GetCell(address)?.DisplayText ?? string.Empty;
                            if (budget.ElapsedMilliseconds >= 4) break;
                        }
                    }, DispatcherPriority.Background);
                    if (more) await Task.Delay(1, token).ConfigureAwait(false);
                }
                (ChartData data, bool unchanged) = await BackgroundWork.Run(() =>
                {
                    (List<string> categories, List<ChartSeriesData> series) = ExtractChartData(
                        chart, address => { token.ThrowIfCancellationRequested(); return values.GetValueOrDefault(address) ?? string.Empty; }, rows, columns);
                    ChartData data = new(categories, series);
                    return (data, previous is not null && SameChartData(previous, data));
                }, token).ConfigureAwait(false);
                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    token.ThrowIfCancellationRequested();
                    if (!ReferenceEquals(this.sheet, model) || !model.Charts.Contains(chart)) return;
                    if (this.chartDrag is not null)
                    {
                        this.chartRefreshDeferred = true;
                        return;
                    }
                    this.BuildChartCard(chart, data, unchanged);
                }, DispatcherPriority.Background);
                // Never construct every chart in a single dispatcher callback.
                await Task.Delay(1, token).ConfigureAwait(false);
            }
        }
        catch (OperationCanceledException) { }
        catch (Exception exception) { Trace.TraceError("Spreadsheet chart refresh failed: {0}", exception); }
    }

    private void UpdateChartSelection()
    {
        foreach (var entry in this.chartCards)
        {
            bool selected = ReferenceEquals(entry.Card.Tag, this.selectedChart);
            entry.Card.BorderBrush = selected ? SelectionBorder : SheetGridLine;
            entry.Card.BorderThickness = new Thickness(selected ? 2 : 1);
        }
    }

    private sealed record ChartDrag(SpreadsheetViewModel Sheet, SheetChart Model, Border Card, Control Chart,
        IPointer Pointer, bool Resize, Point Start, double Left, double Top, double Width, double Height)
    {
        internal bool Moved { get; set; }
    }
    private ChartDrag? chartDrag;

    private void InitializeChartDragging()
    {
        this.PointerMoved += (_, e) =>
        {
            if (this.chartDrag is not { } drag || !ReferenceEquals(e.Pointer.Captured, this)) return;
            this.UpdateChartDrag(drag, e.GetPosition(this.cells));
            e.Handled = true;
        };
        this.PointerReleased += (_, e) =>
        {
            if (this.chartDrag is not { } drag || !ReferenceEquals(e.Pointer.Captured, this) ||
                e.InitialPressMouseButton != MouseButton.Left) return;
            this.UpdateChartDrag(drag, e.GetPosition(this.cells));
            this.FinishChartDrag();
            e.Handled = true;
        };
        this.PointerCaptureLost += (_, _) =>
        {
            if (this.chartDrag is { } drag && !ReferenceEquals(drag.Pointer.Captured, this)) this.FinishChartDrag(cancel: true);
        };
        this.DetachedFromVisualTree += (_, _) =>
        {
            this.FinishChartDrag(cancel: true);
            this.CancelChartRefresh();
        };
    }

    private void BeginChartDrag(PointerPressedEventArgs e, SheetChart chart, Border card, Control control, bool resize)
    {
        if (this.ReadOnly || this.sheet is null || this.chartDrag is not null ||
            !this.sheet.Charts.Contains(chart) || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed) return;
        this.SuspendEditor();
        this.CancelFillDrag();
        this.FinishHeaderResize();
        this.SelectChart(chart);
        this.chartDrag = new ChartDrag(this.sheet, chart, card, control, e.Pointer, resize,
            e.GetPosition(this.cells), Canvas.GetLeft(card), Canvas.GetTop(card), card.Width, card.Height);
        this.chartDragging = true;
        // The grid survives viewport/chart refreshes; a transient chart card does not.
        e.Pointer.Capture(this);
        e.Handled = true;
    }

    private void UpdateChartDrag(ChartDrag drag, Point position)
    {
        double dx = position.X - drag.Start.X, dy = position.Y - drag.Start.Y;
        if (!drag.Moved && Math.Abs(dx) < 4 && Math.Abs(dy) < 4) return;
        drag.Moved = true;
        if (drag.Resize)
        {
            drag.Card.Width = drag.Chart.Width = Math.Clamp(drag.Width + dx / this.zoom, 120, 2400);
            drag.Card.Height = drag.Chart.Height = Math.Clamp(drag.Height + dy / this.zoom, 80, 1600);
        }
        else
        {
            Canvas.SetLeft(drag.Card, this.OffsetToColumn(this.NearestColumn(Math.Max(0, drag.Left + dx))));
            Canvas.SetTop(drag.Card, this.OffsetToRow(this.NearestRow(Math.Max(0, drag.Top + dy))));
        }
    }

    private void FinishChartDrag(bool cancel = false)
    {
        if (this.chartDrag is not { } drag) return;
        this.chartDrag = null;
        this.chartDragging = false;
        if (ReferenceEquals(drag.Pointer.Captured, this)) drag.Pointer.Capture(null);
        if (cancel)
        {
            Canvas.SetLeft(drag.Card, drag.Left);
            Canvas.SetTop(drag.Card, drag.Top);
            drag.Card.Width = drag.Chart.Width = drag.Width;
            drag.Card.Height = drag.Chart.Height = drag.Height;
        }
        else if (drag.Moved && ReferenceEquals(this.sheet, drag.Sheet) && drag.Sheet.Charts.Contains(drag.Model))
        {
            if (drag.Resize) drag.Sheet.ResizeChart(drag.Model, drag.Card.Width, drag.Card.Height);
            else drag.Sheet.MoveChart(drag.Model, this.NearestRow(Canvas.GetTop(drag.Card)),
                this.NearestColumn(Canvas.GetLeft(drag.Card)), drag.Card.Width, drag.Card.Height);
        }
        if (this.chartRefreshDeferred)
        {
            this.chartRefreshDeferred = false;
            this.ScheduleChartRefresh();
        }
        this.AdjustRowsForScroll();
        this.AdjustColumnsForScroll();
    }
}
