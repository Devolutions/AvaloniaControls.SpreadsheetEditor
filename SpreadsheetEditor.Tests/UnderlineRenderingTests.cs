using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.VisualTree;
using SkiaSharp;
using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;
using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class UnderlineRenderingTests
{
    [AvaloniaTheory]
    [InlineData("Calibri", 14.6666667, 1.0, false)]
    [InlineData("Arial", 14.6666667, 1.0, false)]
    [InlineData("Calibri", 24, 1.0, false)]
    [InlineData("Calibri", 14.6666667, 2.0, false)]
    [InlineData("Arial", 28, 1.0, true)]
    [InlineData("Calibri", 28, 1.0, true)]
    public async Task DoubleUnderlinePaintsTwoSeparatedLines(string family, double fontSize, double zoom, bool automaticHeight)
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[] { CellData.Plain("HHHHH") with
            {
                Underline = true, FontSize = fontSize, FontFamily = family,
                Appearance = new CellAppearance { DoubleUnderline = true },
            } } }, ColumnWidths = new[] { 350d }, RowHeights = automaticHeight ? Array.Empty<double>() : new[] { 60d },
        });
        editor.Zoom = zoom;
        Window window = new() { Content = editor, Width = 1000, Height = 500 }; window.Show();
        try
        {
            window.UpdateLayout(); window.CaptureRenderedFrame();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
            TextBlock text = editor.GetVisualDescendants().OfType<TextBlock>().Single(t => t.Name == "CellText" && Equals(t.Tag, new CellAddress(0, 0)));
            int width = (int)Math.Ceiling(text.TextLayout.Width), height = (int)Math.Ceiling(text.Bounds.Height + 12);
            using RenderTargetBitmap bitmap = new(new PixelSize(width, height), new Vector(96, 96));
            bitmap.Render(text);
            using MemoryStream stream = new(); bitmap.Save(stream, new PngBitmapEncoderOptions());
            using SKBitmap pixels = SKBitmap.Decode(stream.ToArray());
            int bands = 0;
            bool previous = false;
            for (int y = (int)Math.Ceiling(text.TextLayout.TextLines[0].Baseline); y < height; y++)
            {
                int dark = Enumerable.Range(1, width - 2).Count(x =>
                {
                    SKColor pixel = pixels.GetPixel(x, y);
                    return pixel.Alpha > 80 && pixel.Red < 80 && pixel.Green < 80 && pixel.Blue < 80;
                });
                bool line = dark > (width - 2) * 0.8;
                if (line && !previous) bands++;
                previous = line;
            }
            Assert.Equal(2, bands);
        }
        finally { editor.NewSheet(1, 1); window.Close(); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle); }
    }

    [AvaloniaFact]
    public void TurningUnderlineOffDoesNotRestoreDoubleUnderlineOnSave()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData { Cells = new[] { new[] { CellData.Plain("Text") with
            { Underline = true, Appearance = new CellAppearance { DoubleUnderline = true } } } } });
        editor.FindControl<SpreadsheetGrid>("Grid")!.ToggleUnderline();
        using MemoryStream saved = new(); editor.SaveXlsx(saved); saved.Position = 0;
        var cell = XlsxFile.Read(saved).Cells[0][0];
        Assert.False(cell.Underline);
        Assert.False(cell.Appearance!.DoubleUnderline);
    }
}
