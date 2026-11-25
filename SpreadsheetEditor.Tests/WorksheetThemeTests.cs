namespace SpreadsheetEditor.Tests;

using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Documents.TextModel;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SkiaSharp;
using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

public sealed class WorksheetThemeTests
{
    [AvaloniaTheory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task OnlyOptedInDarkWorksheetsChangeTheirAutomaticColors(bool enabled, bool dark)
    {
        EditorControl editor = CreateEditor();
        Assert.False(editor.MatchWorksheetThemeToUI);
        editor.MatchWorksheetThemeToUI = enabled;
        int changes = 0;
        editor.ContentChanged += (_, _) => changes++;
        Window window = new() { Content = editor, Width = 900, Height = 450,
            RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light };
        window.Show();
        try
        {
            await Frame(window);
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            MenuFlyout menu = Assert.IsType<MenuFlyout>(editor.FindControl<Button>("OverflowButton")!.Flyout);
            MenuItem option = menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, EditorText.Get(EditorString.MatchWorksheetThemeToUI)));
            Assert.Equal(dark, option.IsVisible);
            bool darkSheet = enabled && dark;
            Assert.Equal(darkSheet ? LightInk : DarkInk, Ink(CellText(grid, 0, 0).Foreground));
            Assert.Equal(darkSheet ? DarkSurface : Colors.White, Surface(grid));
            Assert.True(AutomaticInkPixels(grid, darkSheet) > 0);
            TextBlock colored = CellText(grid, 0, 1);
            Assert.Equal(Colors.Red, Ink(colored.Inlines![0].Foreground));
            Assert.Equal(Colors.White, Ink(CellText(grid, 0, 2).Inlines![0].Foreground));
            Assert.Equal(Colors.Black, Ink(CellText(grid, 0, 3).Inlines![0].Foreground));
            Assert.Equal(Colors.Gray, editor.ViewModel.Rows[0].Cells[1].BackgroundColor);
            Assert.Equal(0, changes);
            Assert.False(editor.CanUndo);

            using MemoryStream saved = new();
            await editor.SaveXlsxAsync(saved);
            saved.Position = 0;
            CellData[] cells = XlsxFile.ReadWorkbook(saved)[0].Data.Cells[0].Take(3).ToArray();
            Assert.Null(cells[0].Foreground);
            Assert.Equal(Colors.Red, cells[1].Foreground);
            Assert.Equal(Colors.Gray, cells[1].Background);
            Assert.Equal(Colors.White, cells[2].Foreground);
            Assert.Equal(0, changes);
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task RuntimeOptionKeepsCellsAndTheirEditingDocumentWhileTheThemeChanges()
    {
        EditorControl editor = CreateEditor();
        Window window = new() { Content = editor, Width = 900, Height = 450, RequestedThemeVariant = ThemeVariant.Dark };
        window.Show();
        try
        {
            await Frame(window);
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            MenuFlyout menu = Assert.IsType<MenuFlyout>(editor.FindControl<Button>("OverflowButton")!.Flyout);
            MenuItem option = menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, EditorText.Get(EditorString.MatchWorksheetThemeToUI)));
            Separator separator = Assert.IsType<Separator>(menu.Items[menu.Items.IndexOf(option) - 1]);
            Assert.True(option.IsVisible);
            Assert.True(separator.IsVisible);
            TextBlock retained = CellText(grid, 2, 1);
            CellEditor input = grid.GetVisualDescendants().OfType<CellEditor>().Single();
            input.Load(editor.ViewModel.Rows[0].Cells[0]);
            RichTextEditor? rich = input.Child as RichTextEditor;
            FlowDocument? document = rich?.Document;
            if (document is not null)
            {
                document.InitializeTextDocument();
                TextRange range = new(document.ContentStart!, document.ContentEnd!);
                Assert.Same(document.GetValue(RichTextElement.ForegroundProperty), range.GetPropertyValue(RichTextElement.ForegroundProperty));
            }
            editor.MatchWorksheetThemeToUI = true;
            await Frame(window);
            Assert.Same(retained, CellText(grid, 2, 1));
            Assert.Equal(DarkSurface, Surface(grid));
            Assert.Equal(LightInk, Ink(CellText(grid, 0, 0).Foreground));
            Assert.True(AutomaticInkPixels(grid, light: true) > 0);
            if (document is not null)
            {
                Assert.Same(document, rich!.Document);
                Assert.Equal(LightInk, Ink(document.GetValue(RichTextElement.ForegroundProperty)));
                Assert.Equal(LightInk, Ink(rich.CaretBrush));
            }
            editor.MatchWorksheetThemeToUI = false;
            await Frame(window);
            Assert.Same(retained, CellText(grid, 2, 1));
            Assert.Equal(Colors.White, Surface(grid));
            editor.MatchWorksheetThemeToUI = true;
            window.RequestedThemeVariant = ThemeVariant.Light;
            await Frame(window);
            Assert.Equal(Colors.White, Surface(grid));
            Assert.Equal(DarkInk, Ink(CellText(grid, 0, 0).Foreground));
            Assert.False(option.IsVisible);
            Assert.False(separator.IsVisible);
            Assert.True(editor.MatchWorksheetThemeToUI);
            Assert.True(option.IsChecked);
            window.RequestedThemeVariant = ThemeVariant.Dark;
            await Frame(window);
            Assert.True(option.IsVisible);
            Assert.True(separator.IsVisible);
            Assert.True(option.IsChecked);
            Assert.Equal(DarkSurface, Surface(grid));
            Assert.Null(editor.ViewModel.Rows[0].Cells[0].ForegroundColor);
            if (rich is not null)
            {
                input.Load(editor.ViewModel.Rows[0].Cells[2]);
                FlowDocument explicitDocument = rich.Document!;
                explicitDocument.InitializeTextDocument();
                TextRange explicitRange = new(explicitDocument.ContentStart!, explicitDocument.ContentEnd!);
                IBrush? explicitBrush = explicitRange.GetPropertyValue(RichTextElement.ForegroundProperty);
                Assert.NotSame(explicitDocument.GetValue(RichTextElement.ForegroundProperty), explicitBrush);
                editor.MatchWorksheetThemeToUI = false;
                Assert.Equal(Colors.White, Ink(explicitBrush));
            }
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task WorksheetPaletteIsIndependentForEachEditorAndAvailableInReadOnlyMode()
    {
        EditorControl first = CreateEditor();
        EditorControl second = CreateEditor();
        first.MatchWorksheetThemeToUI = true;
        first.IsReadOnly = true;
        Grid content = new() { ColumnDefinitions = new ColumnDefinitions("*,*") };
        Grid.SetColumn(second, 1);
        content.Children.Add(first);
        content.Children.Add(second);
        Window window = new() { Content = content, Width = 1200, Height = 450, RequestedThemeVariant = ThemeVariant.Dark };
        window.Show();
        try
        {
            await Frame(window);
            Assert.Equal(DarkSurface, Surface(first.FindControl<SpreadsheetGrid>("Grid")!));
            Assert.Equal(Colors.White, Surface(second.FindControl<SpreadsheetGrid>("Grid")!));
            MenuFlyout menu = Assert.IsType<MenuFlyout>(first.FindControl<Button>("OverflowButton")!.Flyout);
            MenuItem option = menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, EditorText.Get(EditorString.MatchWorksheetThemeToUI)));
            Assert.True(option.IsEnabled);
            Assert.True(option.IsChecked);
            option.SetCurrentValue(MenuItem.IsCheckedProperty, false);
            Assert.False(first.MatchWorksheetThemeToUI);
            first.MatchWorksheetThemeToUI = true;
            Assert.True(option.IsChecked);
            Assert.False(first.CanUndo);
        }
        finally
        {
            second.NewSheet(1, 1);
            await Close(first, window);
        }
    }

    private static readonly Color DarkSurface = Color.FromRgb(0x20, 0x20, 0x20);
    private static readonly Color DarkInk = Color.FromRgb(0x20, 0x20, 0x20);
    private static readonly Color LightInk = Color.FromRgb(0xF3, 0xF3, 0xF3);

    private static EditorControl CreateEditor()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[]
            {
                new[] { CellData.Plain("Automatic"), CellData.Plain("Red") with { Foreground = Colors.Red, Background = Colors.Gray },
                    CellData.Plain("White") with { Foreground = Colors.White }, CellData.Plain("Black") with { Foreground = Colors.Black } },
                Enumerable.Repeat(CellData.Plain(string.Empty), 4).ToArray(),
                Enumerable.Repeat(CellData.Plain(string.Empty), 4).ToArray(),
            },
        });
        return editor;
    }

    private static Color Ink(IBrush? brush) => Assert.IsAssignableFrom<ISolidColorBrush>(brush).Color;

    private static TextBlock CellText(SpreadsheetGrid grid, int row, int column) => grid.GetVisualDescendants()
        .OfType<TextBlock>().Single(text => text.Name == "CellText" && Equals(text.Tag, new CellAddress(row, column)));

    private static Color Surface(SpreadsheetGrid grid)
    {
        TextBlock blank = CellText(grid, 2, 1);
        Point point = blank.TranslatePoint(new Point(10, 4), grid)!.Value;
        using SKBitmap pixels = Render(grid);
        SKColor color = pixels.GetPixel((int)point.X, (int)point.Y);
        return Color.FromArgb(color.Alpha, color.Red, color.Green, color.Blue);
    }

    private static int AutomaticInkPixels(SpreadsheetGrid grid, bool light)
    {
        TextBlock text = CellText(grid, 0, 0);
        Point point = text.TranslatePoint(new Point(1, 1), grid)!.Value;
        using SKBitmap pixels = Render(grid);
        int count = 0;
        for (int y = (int)point.Y; y < point.Y + text.Bounds.Height - 2; y++)
        {
            for (int x = (int)point.X; x < point.X + text.Bounds.Width - 2; x++)
            {
                SKColor pixel = pixels.GetPixel(x, y);
                if (light ? pixel.Red > 180 && pixel.Green > 180 && pixel.Blue > 180
                    : pixel.Red < 100 && pixel.Green < 100 && pixel.Blue < 100)
                {
                    count++;
                }
            }
        }
        return count;
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
