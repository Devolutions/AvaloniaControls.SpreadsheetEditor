using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Models;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class ChartInteractionTests
{
    [AvaloniaTheory]
    [InlineData(false, 1.0)]
    [InlineData(true, 1.0)]
    [InlineData(false, 0.5)]
    [InlineData(true, 2.0)]
    public async Task ChartDragsSurviveRepaintsWithoutShrinkingTheSheet(bool resize, double zoom)
    {
        EditorControl editor = CreateEditor();
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Window window = new() { Content = editor, Width = 1400, Height = 1000 };
        window.Show();
        grid.Zoom = zoom;
        try
        {
            Border card = await WaitForCard(grid, window);
            int rows = editor.ViewModel.RowCount, columns = editor.ViewModel.ColumnCount;
            SheetChart chart = editor.ViewModel.Charts[0];
            Control grab = resize ? card.GetVisualDescendants().OfType<Border>().Single(b => b.Name == "ChartResizeGrip") : card;
            Point start = grab.TranslatePoint(resize ? new Point(7, 7) : new Point(20, 20), window)!.Value;
            double originalWidth = card.Width, originalHeight = card.Height;
            window.MouseDown(start, MouseButton.Left);
            for (int step = 1; step <= 3; step++)
            {
                window.MouseMove(new Point(start.X + 30 * step * zoom, start.Y + 24 * step * zoom));
                grid.Rebuild(); // layout/data changes used to destroy the captured chart card
                await RenderFrame(window);
                Assert.Same(card, FindCard(grid));
            }
            Point end = new(start.X + 90 * zoom, start.Y + 72 * zoom);
            window.MouseUp(end, MouseButton.Left);
            await Task.Delay(50);
            await RenderFrame(window);
            Assert.Same(card, FindCard(grid));
            Assert.Equal(rows, editor.ViewModel.RowCount);
            Assert.Equal(columns, editor.ViewModel.ColumnCount);
            if (resize)
            {
                Assert.Equal(originalWidth + 90, chart.Width, precision: 4);
                Assert.Equal(originalHeight + 72, chart.Height, precision: 4);
            }
            else
            {
                Assert.Equal(3, chart.AnchorColumn);
                Assert.Equal(9, chart.AnchorRow);
            }
        }
        finally { await Cleanup(editor, window); }
    }

    [AvaloniaFact]
    public async Task ChartGearOpensOneDialogAndKeepsBlankWorksheetArea()
    {
        EditorControl editor = CreateEditor();
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Window window = new() { Content = editor, Width = 900, Height = 650 };
        window.Show();
        try
        {
            Border card = await WaitForCard(grid, window);
            int rows = editor.ViewModel.RowCount, columns = editor.ViewModel.ColumnCount;
            Button gear = card.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "EditChartButton");
            Point point = gear.TranslatePoint(new Point(11, 11), window)!.Value;
            window.MouseDown(point, MouseButton.Left);
            window.MouseUp(point, MouseButton.Left);
            gear.RaiseEvent(new RoutedEventArgs(Button.ClickEvent)); // duplicate host/routed request
            await RenderFrame(window);
            Window dialog = Assert.Single(window.OwnedWindows);
            Assert.Same(card, FindCard(grid));
            Assert.Equal(rows, editor.ViewModel.RowCount);
            Assert.Equal(columns, editor.ViewModel.ColumnCount);
            dialog.Close(new ChartBuilderResult(ChartKind.Line, "A1:B3", "Edited chart", false, true, true, false,
                AxisNumberFormat.Number, null, null, null));
            await Task.Delay(50);
            await RenderFrame(window);
            Assert.Empty(window.OwnedWindows);
            Assert.Equal("Edited chart", editor.ViewModel.Charts[0].Title);
            Assert.Equal(rows, editor.ViewModel.RowCount);
            Assert.Equal(columns, editor.ViewModel.ColumnCount);
        }
        finally { await Cleanup(editor, window); }
    }

    [AvaloniaFact]
    public async Task EscapeCancelsChartDragAndAllowsAnotherDrag()
    {
        EditorControl editor = CreateEditor();
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Window window = new() { Content = editor, Width = 900, Height = 650 };
        window.Show();
        try
        {
            Border card = await WaitForCard(grid, window);
            Point start = card.TranslatePoint(new Point(20, 20), window)!.Value;
            Point end = new(start.X + 90, start.Y + 72);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, null);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal(2, editor.ViewModel.Charts[0].AnchorColumn);
            Assert.Equal(6, editor.ViewModel.Charts[0].AnchorRow);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal(3, editor.ViewModel.Charts[0].AnchorColumn);
            Assert.Equal(9, editor.ViewModel.Charts[0].AnchorRow);
        }
        finally { await Cleanup(editor, window); }
    }

    [AvaloniaFact]
    public async Task ReadOnlyBlocksChartEditingAndMoving()
    {
        EditorControl editor = CreateEditor();
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Window window = new() { Content = editor, Width = 900, Height = 650 };
        window.Show();
        try
        {
            Border card = await WaitForCard(grid, window);
            grid.ReadOnly = true; // also used while a lazy worksheet preview is being replaced
            Button gear = card.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "EditChartButton");
            Assert.False(gear.IsEnabled);
            gear.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Point start = card.TranslatePoint(new Point(20, 20), window)!.Value;
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(new Point(start.X + 100, start.Y + 100));
            window.MouseUp(new Point(start.X + 100, start.Y + 100), MouseButton.Left);
            await RenderFrame(window);
            Assert.Empty(window.OwnedWindows);
            Assert.Equal(2, editor.ViewModel.Charts[0].AnchorColumn);
            Assert.Equal(6, editor.ViewModel.Charts[0].AnchorRow);
        }
        finally { await Cleanup(editor, window); }
    }

    private static EditorControl CreateEditor()
    {
        CellData[][] cells = Enumerable.Range(0, 100)
            .Select(_ => Enumerable.Repeat(CellData.Plain(string.Empty), 40).ToArray()).ToArray();
        for (int row = 0; row < 3; row++)
        {
            cells[row][0] = CellData.Plain($"Label {row + 1}");
            cells[row][1] = CellData.Plain((row + 1).ToString());
        }
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = cells,
            RowHeights = Enumerable.Repeat(24d, 100).ToArray(),
            ColumnWidths = Enumerable.Repeat(90d, 40).ToArray(),
            Charts = new[] { new SheetChart
            {
                Kind = ChartKind.Line,
                Series = new[] { new ChartSeriesSpec("Values", new CellRange(0, 1, 3, 1)) },
                Categories = new CellRange(0, 0, 3, 1),
                AnchorRow = 6, AnchorColumn = 2, Width = 300, Height = 180,
            } },
        });
        return editor;
    }

    private static Border? FindCard(SpreadsheetGrid grid) => grid.GetVisualDescendants().OfType<Border>()
        .SingleOrDefault(border => border.Name == "ChartCard");

    private static async Task<Border> WaitForCard(SpreadsheetGrid grid, Window window)
    {
        Stopwatch timeout = Stopwatch.StartNew();
        while (FindCard(grid) is null && timeout.Elapsed < TimeSpan.FromSeconds(5))
        {
            await Task.Delay(10);
            await RenderFrame(window);
        }
        return Assert.IsType<Border>(FindCard(grid));
    }

    private static async Task RenderFrame(Window window)
    {
        await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
        window.UpdateLayout();
    }

    private static async Task Cleanup(EditorControl editor, Window window)
    {
        foreach (Window owned in window.OwnedWindows.ToArray()) owned.Close();
        editor.NewSheet(1, 1);
        window.Close();
        await Task.Delay(30);
        await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
    }
}
