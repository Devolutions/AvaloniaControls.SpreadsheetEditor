namespace SpreadsheetEditor.Tests;

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Interactivity;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using SkiaSharp;
using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;
using Xunit;

public sealed class BorderMenuRenderingTests
{
    [AvaloniaTheory]
    [InlineData("thin", 1d, false, 1, 1)]
    [InlineData("thin", 1.5d, false, 1, 1)]
    [InlineData("medium", 1d, false, 1, 1)]
    [InlineData("double", 1d, false, 2, 1)]
    [InlineData("thin", 1d, true, 1, 2)]
    [InlineData("double", 1d, true, 2, 2)]
    [InlineData("thick", 1d, false, 1, 3)]
    [InlineData("thick", 1.5d, false, 1, 3)]
    public async Task SharedBottomBorderDoesNotPaintTwiceOrLeaveAGridline(string style, double zoom, bool merged, int expectedBands, int span)
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = Enumerable.Range(0, 4).Select(_ => Enumerable.Range(0, 4)
                .Select(_ => CellData.Plain("") with { Appearance = new CellAppearance() }).ToArray()).ToArray(),
            ColumnWidths = new[] { 100d, 100d, 100d, 100d },
            RowHeights = new[] { 30d, 30d, 30d, 30d },
        });
        if (merged) editor.ViewModel.MergeCells(new CellRange(0, 0, span, span));
        editor.ViewModel.ApplyBorders(new CellRange(0, 0, span, span), CellBorderPreset.Outside, new CellEdge(style, Colors.Red));
        editor.Zoom = zoom;
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        grid.Rebuild();
        grid.SelectCell(new CellAddress(3, 3));
        Window window = new() { Content = editor, Width = 1000, Height = 500 };
        window.Show();
        try
        {
            await Frame(window);
            Border cell = grid.GetVisualDescendants().OfType<Border>().Single(border => Equals(border.Tag, new CellAddress(merged ? 0 : span - 1, 0)));
            Point boundary = cell.TranslatePoint(new Point(cell.Bounds.Width / 2 - 10, cell.Bounds.Height), grid)!.Value;
            using SKBitmap pixels = Render(grid);
            bool IsRed(int y)
            {
                SKColor pixel = pixels.GetPixel((int)boundary.X, y);
                return pixel.Alpha > 80 && pixel.Red > 160 && pixel.Green < 130 && pixel.Blue < 130;
            }
            int bands = 0;
            bool previous = false;
            for (int y = (int)boundary.Y - 9; y < (int)boundary.Y + 6; y++)
            {
                bool red = IsRed(y);
                if (red && !previous) bands++;
                previous = red;
            }
            Assert.Equal(expectedBands, bands);
            Assert.True(IsRed((int)Math.Floor(boundary.Y) - 1), "The custom border must cover the grey bottom gridline.");
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task BorderMenuShowsIconsLineSamplesAndTheMarkdownColorPicker()
    {
        EditorControl editor = new();
        Window window = new() { Content = editor, Width = 1100, Height = 600 };
        window.Show();
        try
        {
            await Frame(window);
            Button button = editor.FindControl<Button>("BordersButton")!;
            MenuFlyout menu = Assert.IsType<MenuFlyout>(button.Flyout);
            MenuItem[] items = menu.Items.OfType<MenuItem>().ToArray();
            Assert.All(items.Take(8), item => Assert.IsType<Canvas>(item.Icon));
            menu.ShowAt(button);
            await Frame(window);
            MenuItem color = items.Single(item => Equals(item.Header, EditorText.Get(EditorString.BorderColor)));
            color.IsSubMenuOpen = true;
            await Frame(window);
            ColorView picker = Assert.IsType<ColorView>(Assert.IsType<MenuItem>(Assert.Single(color.Items)).Header);
            Assert.False(picker.IsAlphaEnabled);
            Assert.Contains(Colors.Black, picker.PaletteColors!);
            Assert.Contains(picker.GetVisualDescendants(), control => control is ColorSpectrum && control.Bounds.Width > 0);
            picker.Color = Colors.Purple;
            Assert.Equal(Colors.Purple, editor.FindControl<SpreadsheetGrid>("Grid")!.BorderColor);
            color.IsSubMenuOpen = false;
            MenuItem styles = items.Single(item => Equals(item.Header, EditorText.Get(EditorString.BorderStyle)));
            styles.IsSubMenuOpen = true;
            await Frame(window);
            foreach (MenuItem item in styles.Items.OfType<MenuItem>())
            {
                Canvas sample = Assert.IsType<Canvas>(item.Header);
                Assert.All(sample.Children.OfType<Line>(), line => Assert.NotNull(line.Stroke));
            }
            styles.IsSubMenuOpen = false;
            menu.Hide();
            MenuFlyout more = Assert.IsType<MenuFlyout>(editor.FindControl<Button>("OverflowButton")!.Flyout);
            Assert.DoesNotContain(more.Items.OfType<MenuItem>(), item => Equals(item.Header, EditorText.Get(EditorString.Borders))
                || Equals(item.Header, EditorText.Get(EditorString.MergeCells)));
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task ItalicGlyphHasRoomForItsSlantedInk()
    {
        EditorControl editor = new();
        Window window = new() { Content = editor, Width = 1100, Height = 450 };
        window.Show();
        try
        {
            await Frame(window);
            Button glyph = editor.FindControl<Button>("ItalicButton")!;
            Assert.IsType<PathIcon>(glyph.Content);
            using SKBitmap pixels = Render(glyph);
            int[] inkColumns = Enumerable.Range(0, pixels.Width).Where(x => Enumerable.Range(0, pixels.Height).Any(y =>
            {
                SKColor pixel = pixels.GetPixel(x, y);
                return pixel.Alpha > 60 && pixel.Red < 180 && pixel.Green < 180 && pixel.Blue < 180;
            })).ToArray();
            Assert.NotEmpty(inkColumns);
            Assert.True(inkColumns[0] > 0 && inkColumns[^1] < pixels.Width - 1);
            bool IsInk(SKBitmap bitmap, int x, int y)
            {
                SKColor pixel = bitmap.GetPixel(x, y);
                return pixel.Alpha > 60 && pixel.Red < 180 && pixel.Green < 180 && pixel.Blue < 180;
            }
            int[] inkRows = Enumerable.Range(0, pixels.Height)
                .Where(y => Enumerable.Range(0, pixels.Width).Any(x => IsInk(pixels, x, y))).ToArray();
            int RowWidth(int y) => Enumerable.Range(0, pixels.Width).Count(x => IsInk(pixels, x, y));
            int middleWidth = RowWidth(inkRows[inkRows.Length / 2]);
            Assert.True(RowWidth(inkRows[0]) > middleWidth, "The italic I needs a top stroke.");
            Assert.True(RowWidth(inkRows[^1]) > middleWidth, "The italic I needs a bottom stroke.");
            using SKBitmap bold = Render(editor.FindControl<Button>("BoldButton")!);
            int boldHeight = Enumerable.Range(0, bold.Height)
                .Count(y => Enumerable.Range(0, bold.Width).Any(x => IsInk(bold, x, y)));
            Assert.InRange(inkRows.Length, boldHeight - 2, boldHeight + 2);
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task ReselectingAnOutlineAllowsEditingItsColorAndThicknessWithoutApplyingAnotherPreset()
    {
        EditorControl editor = new();
        editor.NewSheet(6, 6);
        CellRange range = new(1, 1, 2, 2);
        editor.ViewModel.ApplyBorders(range, CellBorderPreset.Outside, new CellEdge("thick", Colors.Red));
        Window window = new() { Content = editor, Width = 1100, Height = 600 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            grid.SelectCell(new CellAddress(4, 4));
            grid.SelectCell(new CellAddress(1, 1));
            grid.FocusForSelection();
            window.KeyPress(Key.Right, RawInputModifiers.Shift, PhysicalKey.ArrowRight, null);
            window.KeyPress(Key.Down, RawInputModifiers.Shift, PhysicalKey.ArrowDown, null);
            Assert.Equal(range, grid.SelectionRange);
            await Frame(window);
            Button button = editor.FindControl<Button>("BordersButton")!;
            MenuFlyout menu = Assert.IsType<MenuFlyout>(button.Flyout);
            menu.ShowAt(button);
            await Frame(window);
            Assert.Equal(Colors.Red, grid.BorderColor);
            Assert.Equal("thick", grid.BorderLineStyle);
            MenuItem color = menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, EditorText.Get(EditorString.BorderColor)));
            color.IsSubMenuOpen = true;
            await Frame(window);
            ColorView picker = Assert.IsType<ColorView>(Assert.IsType<MenuItem>(Assert.Single(color.Items)).Header);
            Assert.Equal(Colors.Red, picker.Color);
            picker.Color = Colors.Blue;
            await Frame(window);
            picker.Color = Colors.Green;
            color.IsSubMenuOpen = false;
            await Frame(window);
            Assert.Equal(new CellEdge("thick", Colors.Green), editor.ViewModel.Rows[1].Cells[1].Appearance!.Left);
            Assert.Null(editor.ViewModel.Rows[1].Cells[1].Appearance!.Right);
            MenuItem styles = menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, EditorText.Get(EditorString.BorderStyle)));
            styles.IsSubMenuOpen = true;
            await Frame(window);
            MenuItem dotted = styles.Items.OfType<MenuItem>().Single(item => Equals(item.Tag, "dotted"));
            dotted.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            Assert.Equal(new CellEdge("dotted", Colors.Green), editor.ViewModel.Rows[2].Cells[2].Appearance!.Bottom);
            styles.IsSubMenuOpen = false;
            menu.Hide();
            editor.ViewModel.Undo();
            Assert.Equal(new CellEdge("thick", Colors.Green), editor.ViewModel.Rows[1].Cells[1].Appearance!.Left);
            editor.ViewModel.Undo();
            Assert.Equal(new CellEdge("thick", Colors.Red), editor.ViewModel.Rows[1].Cells[1].Appearance!.Left);
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task FormatCellsCanRecolorExistingBordersWhileKeepingTheirLayout()
    {
        EditorControl editor = new();
        editor.NewSheet(5, 5);
        editor.ViewModel.ApplyBorders(new CellRange(0, 0, 2, 2), CellBorderPreset.Outside, new CellEdge("medium", Colors.Red));
        Window window = new() { Content = editor, Width = 1100, Height = 650 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            grid.SelectCell(new CellAddress(0, 0));
            grid.FocusForSelection();
            window.KeyPress(Key.Right, RawInputModifiers.Shift, PhysicalKey.ArrowRight, null);
            window.KeyPress(Key.Down, RawInputModifiers.Shift, PhysicalKey.ArrowDown, null);
            await Frame(window);
            MenuFlyout menu = grid.GetVisualDescendants().OfType<Control>()
                .Select(control => control.ContextFlyout).OfType<MenuFlyout>().First(flyout => flyout.Items.OfType<MenuItem>()
                    .Any(item => Equals(item.Header, EditorText.Get(EditorString.FormatCellsMenu))));
            MenuItem format = menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, EditorText.Get(EditorString.FormatCellsMenu)));
            format.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            Window dialog = Assert.Single(window.OwnedWindows);
            dialog.UpdateLayout();
            ColorPicker picker = Assert.Single(dialog.GetVisualDescendants().OfType<ColorPicker>());
            Assert.Equal(Colors.Red, picker.Color);
            picker.Color = Colors.Blue;
            Button apply = dialog.GetVisualDescendants().OfType<Button>().Single(button => button.IsDefault);
            apply.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Task.Yield();
            Assert.Empty(window.OwnedWindows);
            Assert.Equal(new CellEdge("medium", Colors.Blue), editor.ViewModel.Rows[0].Cells[0].Appearance!.Left);
            Assert.Equal(new CellEdge("medium", Colors.Blue), editor.ViewModel.Rows[1].Cells[1].Appearance!.Bottom);
            Assert.Null(editor.ViewModel.Rows[0].Cells[0].Appearance!.Right);
        }
        finally { await Close(editor, window); }
    }

    private static SKBitmap Render(Control control)
    {
        using RenderTargetBitmap image = new(new PixelSize((int)Math.Ceiling(control.Bounds.Width), (int)Math.Ceiling(control.Bounds.Height)), new Vector(96, 96));
        image.Render(control);
        using MemoryStream stream = new();
        image.Save(stream, new PngBitmapEncoderOptions());
        return SKBitmap.Decode(stream.ToArray());
    }

    private static async Task Frame(Window window)
    {
        window.UpdateLayout();
        window.CaptureRenderedFrame();
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        window.UpdateLayout();
        window.CaptureRenderedFrame();
    }

    private static async Task Close(EditorControl editor, Window window)
    {
        editor.NewSheet(1, 1);
        window.Close();
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
    }
}
