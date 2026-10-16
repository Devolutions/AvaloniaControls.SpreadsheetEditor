using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class HeaderResizeTests
{
    [AvaloniaTheory]
    [InlineData(true, 1.0, -2)]
    [InlineData(true, 1.0, 2)]
    [InlineData(false, 1.0, -2)]
    [InlineData(false, 1.0, 2)]
    [InlineData(true, 0.5, 0)]
    [InlineData(false, 2.0, 0)]
    public async Task DividerDragContinuesAcrossRenderedFramesAndPersistsItsFinalSize(bool column, double zoom, double grabOffset)
    {
        EditorControl editor = CreateEditor();
        Window window = new() { Content = editor, Width = 800, Height = 500 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            grid.Zoom = zoom;
            await RenderFrame(window);
            Border handle = FirstHandle(grid, column);
            Point grab = new(handle.Bounds.Width / 2 + (column ? grabOffset : 0),
                handle.Bounds.Height / 2 + (column ? 0 : grabOffset));
            Point start = handle.TranslatePoint(grab, window)!.Value;
            double original = column ? 80 : 24;
            int notifications = 0;
            editor.ContentChanged += (_, _) => notifications++;
            window.MouseDown(start, Avalonia.Input.MouseButton.Left);
            for (int step = 1; step <= 3; step++)
            {
                window.MouseMove(new Point(start.X + (column ? 20 * step : 0), start.Y + (column ? 0 : 20 * step)));
                await RenderFrame(window);
                Assert.Equal(original + 20 * step / zoom, Size(editor, column), precision: 5);
            }
            Point end = new(start.X + (column ? 70 : 0), start.Y + (column ? 0 : 70));
            window.MouseUp(end, Avalonia.Input.MouseButton.Left);
            Assert.Equal(original + 70 / zoom, Size(editor, column), precision: 5);
            Assert.Equal(1, notifications);

            using MemoryStream saved = new();
            await editor.SaveXlsxAsync(saved);
            saved.Position = 0;
            SheetData reopened = XlsxFile.Read(saved);
            double persisted = column ? reopened.ColumnWidths[0] : reopened.RowHeights[0];
            Assert.InRange(persisted, original + 70 / zoom - 1, original + 70 / zoom + 1);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaTheory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task EscapeRestoresTheOriginalDimensionAndAllowsAnotherDrag(bool column)
    {
        EditorControl editor = CreateEditor();
        Window window = new() { Content = editor, Width = 800, Height = 500 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            await RenderFrame(window);
            Point start = GrabPoint(grid, window, column);
            double original = Size(editor, column);
            int notifications = 0;
            editor.ContentChanged += (_, _) => notifications++;
            Point end = new(start.X + (column ? 50 : 0), start.Y + (column ? 0 : 50));
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            await RenderFrame(window);
            Assert.Equal(original + 50, Size(editor, column));
            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, null);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal(original, Size(editor, column));
            Assert.Equal(0, notifications);
            await RenderFrame(window);
            start = GrabPoint(grid, window, column);
            end = new(start.X + (column ? 30 : 0), start.Y + (column ? 0 : 30));
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal(original + 30, Size(editor, column));
            Assert.Equal(1, notifications);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaTheory]
    [InlineData(true, 0.5, 24)]
    [InlineData(false, 0.5, 16)]
    [InlineData(true, 2.0, 24)]
    [InlineData(false, 2.0, 16)]
    public async Task MinimumDimensionsAreLogicalSizesAtAnyZoom(bool column, double zoom, double minimum)
    {
        EditorControl editor = CreateEditor();
        Window window = new() { Content = editor, Width = 800, Height = 500 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            grid.Zoom = zoom;
            await RenderFrame(window);
            Point start = GrabPoint(grid, window, column);
            Point end = new(start.X - (column ? 1000 : 0), start.Y - (column ? 0 : 1000));
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            await RenderFrame(window);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal(minimum, Size(editor, column));
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaFact]
    public async Task ReadOnlyPreventsResizingAndCancelsAnActiveDrag()
    {
        EditorControl editor = CreateEditor();
        Window window = new() { Content = editor, Width = 800, Height = 500 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            await RenderFrame(window);
            Point start = GrabPoint(grid, window, column: true);
            Point end = new(start.X + 40, start.Y);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            Assert.Equal(120, Size(editor, column: true));
            editor.IsReadOnly = true;
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal(80, Size(editor, column: true));
            await RenderFrame(window);
            start = GrabPoint(grid, window, column: false);
            end = new(start.X, start.Y + 40);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal(24, Size(editor, column: false));
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaFact]
    public async Task CancellingAnAutoHeightRowRestoresItsAutomaticHeight()
    {
        EditorControl editor = CreateEditor();
        editor.ViewModel.Rows[0].Height = null;
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        grid.Rebuild();
        Window window = new() { Content = editor, Width = 800, Height = 500 };
        window.Show();
        try
        {
            await RenderFrame(window);
            Point start = GrabPoint(grid, window, column: false);
            Point end = new(start.X, start.Y + 40);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            Assert.NotNull(editor.ViewModel.Rows[0].Height);
            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, null);
            window.MouseUp(end, MouseButton.Left);
            Assert.Null(editor.ViewModel.Rows[0].Height);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    private static Point GrabPoint(SpreadsheetGrid grid, Window window, bool column)
    {
        Border handle = FirstHandle(grid, column);
        return handle.TranslatePoint(new Point(handle.Bounds.Width / 2, handle.Bounds.Height / 2), window)!.Value;
    }

    private static double Size(EditorControl editor, bool column) => column
        ? editor.ViewModel.Columns[0].Width : editor.ViewModel.Rows[0].Height!.Value;

    private static Border FirstHandle(SpreadsheetGrid grid, bool column) => grid.GetVisualDescendants().OfType<Border>()
        .First(border => border.Name == (column ? "ColumnResizeHandle" : "RowResizeHandle"));

    private static async Task RenderFrame(Window window)
    {
        await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
        window.UpdateLayout();
    }

    private static EditorControl CreateEditor()
    {
        EditorControl editor = new();
        CellData[][] cells = Enumerable.Range(0, 30)
            .Select(_ => Enumerable.Repeat(CellData.Plain(string.Empty), 12).ToArray()).ToArray();
        cells[0][0] = CellData.Plain("Value");
        editor.LoadSheet(new SheetData
        {
            Cells = cells,
            ColumnWidths = Enumerable.Repeat(80d, 12).ToArray(),
            RowHeights = Enumerable.Repeat(24d, 30).ToArray(),
        });
        return editor;
    }
}
