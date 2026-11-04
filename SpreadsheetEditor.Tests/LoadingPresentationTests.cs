using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Xml.Linq;
using Avalonia.Controls;
using Avalonia;
using Avalonia.Controls.Charts;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;
using ClosedXML.Excel;
using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using SpreadsheetEditor.Controls.ViewModels;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;
using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class LoadingPresentationTests
{
    [AvaloniaFact]
    public async Task SmallWorksheetAndHeadersAreAnchoredAtTopLeftBeforeBlankAreaGrows()
    {
        SpreadsheetGrid grid = new() { Sheet = new SpreadsheetViewModel(2, 3) };
        // Hold the asynchronous growth while inspecting the narrow sheet's first layout.
        typeof(SpreadsheetGrid).GetField("adjustingColumns", BindingFlags.NonPublic | BindingFlags.Instance)!.SetValue(grid, true);
        typeof(SpreadsheetGrid).GetField("adjustingRows", BindingFlags.NonPublic | BindingFlags.Instance)!.SetValue(grid, true);
        Window window = new() { Width = 1000, Height = 700, Content = grid };
        window.Show();
        try
        {
            window.UpdateLayout(); window.CaptureRenderedFrame();
            Control Layer(string name) => (Control)typeof(SpreadsheetGrid).GetField(name, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(grid)!;
            var corner = Layer("cornerBox");
            Point origin = Layer("cells").TranslatePoint(default, grid)!.Value;
            Point headers = Layer("columnHeaders").TranslatePoint(default, grid)!.Value;
            Point rows = Layer("rowHeaders").TranslatePoint(default, grid)!.Value;
            Assert.Equal(corner.Bounds.Right, origin.X, precision: 4);
            Assert.Equal(corner.Bounds.Bottom, origin.Y, precision: 4);
            Assert.Equal(origin.X, headers.X, precision: 4);
            Assert.Equal(origin.Y, rows.Y, precision: 4);
        }
        finally { grid.Sheet = null; window.Close(); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle); }
    }

    [AvaloniaFact]
    public async Task PreviewLoadsEnoughActualColumnsToCoverNarrowAndHiddenColumns()
    {
        using XLWorkbook workbook = new();
        var sheet = workbook.AddWorksheet("Narrow");
        for (int column = 1; column <= 80; column++)
        {
            sheet.Column(column).Width = 3;
            sheet.Cell(1, column).Value = "Column " + column;
            sheet.Cell(300, column).Value = column;
        }
        sheet.Columns(1, 20).Hide();
        using MemoryStream source = new(); workbook.SaveAs(source); source.Position = 0;
        EditorControl editor = new();
        Window window = new() { Width = 1000, Height = 700, Content = editor };
        window.Show();
        try
        {
            await editor.OpenXlsxAsync(source);
            var preview = editor.ViewModel;
            Assert.True(preview.ColumnCount > 40);
            Assert.Equal("Column 41", preview.Rows[0].Cells[40].Text);
            double[] widths = preview.Columns.Select(column => column.IsHidden ? 0 : column.Width).ToArray();
            Assert.True(widths.Sum() > 1000);
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            await Until(() => !ReferenceEquals(editor.ViewModel, preview) && !grid.ReadOnly);
            Assert.Equal(widths, editor.ViewModel.Columns.Take(widths.Length).Select(column => column.IsHidden ? 0 : column.Width));
        }
        finally
        {
            editor.NewSheet(1, 1); window.Close();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        }
    }

    [AvaloniaFact]
    public async Task ColumnWidthsOutsideUsedCellsAreAppliedBeforeTheFirstFrameAndAfterGrowth()
    {
        using XLWorkbook workbook = new();
        var sheet = workbook.AddWorksheet("Small");
        sheet.Cell("A1").Value = "Small sheet";
        sheet.Column(2).Width = 28;
        sheet.Column(25).Width = 35;
        sheet.Column(26).Hide();
        using MemoryStream source = new(); workbook.SaveAs(source); source.Position = 0;
        EditorControl editor = new();
        Window window = new() { Width = 1000, Height = 700, Content = editor };
        window.Show();
        try
        {
            await editor.OpenXlsxAsync(source);
            var preview = editor.ViewModel;
            Assert.Equal(28 * 7 + 5, preview.Columns[1].Width, precision: 4);
            preview.EnsureSize(preview.RowCount, 30);
            Assert.Equal(35 * 7 + 5, preview.Columns[24].Width, precision: 4);
            Assert.True(preview.Columns[25].IsHidden);
            await Until(() => !ReferenceEquals(editor.ViewModel, preview));
            Assert.Equal(28 * 7 + 5, editor.ViewModel.Columns[1].Width, precision: 4);
            Assert.Equal(35 * 7 + 5, editor.ViewModel.Columns[24].Width, precision: 4);
            Assert.True(editor.ViewModel.Columns[25].IsHidden);
        }
        finally
        {
            editor.NewSheet(1, 1); window.Close();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        }
    }

    [AvaloniaFact]
    public async Task CompletingPreviewRetainsRenderedChartsAndEditsTheCompletedModel()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[]
            {
                new[] { CellData.Plain("One"), CellData.Plain("1") },
                new[] { CellData.Plain("Two"), CellData.Plain("2") },
            },
            Charts = new[] { new SheetChart
            {
                Kind = ChartKind.Line, Title = "Test chart", AnchorRow = 5, AnchorColumn = 0,
                Width = 300, Height = 180, Categories = new CellRange(0, 0, 2, 1),
                Series = new[] { new ChartSeriesSpec("Values", new CellRange(0, 1, 2, 1)) },
            } },
        });
        using MemoryStream source = new(); editor.SaveXlsx(source); source.Position = 0;
        Window window = new() { Width = 1000, Height = 700, Content = editor };
        window.Show();
        try
        {
            await editor.OpenXlsxAsync(source);
            var preview = editor.ViewModel;
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Border? Card() => grid.GetVisualDescendants().OfType<Border>().SingleOrDefault(card => card.Name == "ChartCard");
            await Until(() => Card() is not null);
            Border card = Card()!;
            Assert.True(grid.ReadOnly);
            SheetChart previewChart = preview.Charts[0];
            await Until(() => !ReferenceEquals(editor.ViewModel, preview) && !grid.ReadOnly);
            await Task.Delay(50); // Allow the asynchronous data comparison to finish too.
            Assert.Same(card, Card());
            Assert.NotSame(previewChart, card.Tag);
            Assert.Same(editor.ViewModel.Charts[0], card.Tag);
            Button gear = card.GetVisualDescendants().OfType<Button>().Single(button => button.Name == "EditChartButton");
            gear.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Until(() => window.OwnedWindows.Count == 1);
            window.OwnedWindows.Single().Close(new ChartBuilderResult(ChartKind.Line, "A1:B2", "Edited after loading", false, true, true, false,
                AxisNumberFormat.Number, null, null, null));
            await Until(() => editor.ViewModel.Charts[0].Title == "Edited after loading");
        }
        finally
        {
            foreach (Window owned in window.OwnedWindows.ToArray()) owned.Close();
            editor.NewSheet(1, 1); window.Close();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        }
    }
    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PreviewAndCompletedSheetHaveTheSameVisibleLayoutAndTableStyles(bool selectAnotherTab)
    {
        using MemoryStream source = CreateWorkbook(selectAnotherTab);
        EditorControl editor = new();
        Window window = new() { Width = 1000, Height = 700, Content = editor };
        window.Show();
        try
        {
            await editor.OpenXlsxAsync(source);
            if (selectAnotherTab)
            {
                ((Button)editor.FindControl<StackPanel>("SheetTabs")!.Children[1])
                    .RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                await Until(() => editor.ViewModel.Name == "Styled");
            }

            var preview = editor.ViewModel;
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Assert.True(grid.ReadOnly); // The quick preview, before the full import is published.
            window.UpdateLayout(); window.CaptureRenderedFrame();
            double[] columnWidths = Geometry(grid, "columnWidths").Take(7).ToArray();
            double[] rowHeights = Geometry(grid, "rowHeights").Take(8).ToArray();
            CellData[][] cells = preview.Rows.Take(8).Select(row => row.Cells.Take(7).Select(cell => cell.ToData()).ToArray()).ToArray();
            Assert.True(preview.Columns[5].IsHidden);
            Assert.NotNull(cells[2][0].Background); // Table header styling must be present immediately.
            Assert.Single(preview.Merges);

            await Until(() => !ReferenceEquals(editor.ViewModel, preview) && !grid.ReadOnly);
            window.UpdateLayout(); window.CaptureRenderedFrame();
            Assert.Equal(columnWidths, Geometry(grid, "columnWidths").Take(7).ToArray());
            Assert.Equal(rowHeights, Geometry(grid, "rowHeights").Take(8).ToArray());
            for (int row = 0; row < cells.Length; row++)
                for (int column = 0; column < cells[row].Length; column++)
                    Assert.Equivalent(cells[row][column], editor.ViewModel.Rows[row].Cells[column].ToData());
            Assert.True(editor.ViewModel.Columns[5].IsHidden);
        }
        finally
        {
            editor.NewSheet(1, 1); window.Close();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        }
    }

    [AvaloniaFact]
    public async Task BlankViewportAndHiddenColumnsBeyondUsedCellsSurvivePreviewPromotion()
    {
        using XLWorkbook workbook = new();
        var sheet = workbook.AddWorksheet("Small");
        sheet.Cell("A1").Value = "small sheet";
        sheet.Column(6).Hide();
        using MemoryStream source = new(); workbook.SaveAs(source); source.Position = 0;
        EditorControl editor = new();
        Window window = new() { Width = 1100, Height = 700, Content = editor };
        window.Show();
        try
        {
            await editor.OpenXlsxAsync(source);
            var preview = editor.ViewModel;
            // The viewport can extend past the used range during the lightweight preview.
            preview.EnsureSize(60, 30);
            Assert.True(preview.Columns[5].IsHidden);
            double blankFont = preview.Rows[30].Cells[20].FontSize;
            await Until(() => !ReferenceEquals(editor.ViewModel, preview));
            Assert.True(editor.ViewModel.RowCount >= 60);
            Assert.True(editor.ViewModel.ColumnCount >= 30);
            Assert.Equal(blankFont, editor.ViewModel.Rows[30].Cells[20].FontSize);
            Assert.True(editor.ViewModel.Columns[5].IsHidden);
            editor.ViewModel.Columns[5].IsHidden = false;
            using MemoryStream saved = new(); editor.SaveXlsx(saved); saved.Position = 0;
            Assert.DoesNotContain(5, XlsxFile.Read(saved).HiddenColumns);
        }
        finally
        {
            editor.NewSheet(1, 1); window.Close();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        }
    }

    private static double[] Geometry(SpreadsheetGrid grid, string field) =>
        (double[])typeof(SpreadsheetGrid).GetField(field, BindingFlags.NonPublic | BindingFlags.Instance)!.GetValue(grid)!;

    private static async Task Until(Func<bool> done)
    {
        Stopwatch timeout = Stopwatch.StartNew();
        while (!done() && timeout.Elapsed < TimeSpan.FromSeconds(15)) await Task.Delay(5);
        Assert.True(done());
    }

    private static MemoryStream CreateWorkbook(bool otherTab)
    {
        using XLWorkbook workbook = new();
        if (otherTab) workbook.AddWorksheet("First").Cell("A1").Value = "First sheet";
        var sheet = workbook.AddWorksheet("Styled");
        sheet.Cell("A1").Value = "Merged title";
        sheet.Range("A1:D1").Merge();
        sheet.Cell("A1").Style.Font.FontName = "Arial";
        sheet.Cell("A1").Style.Font.Bold = true;
        sheet.Row(1).Height = 15; // Explicit default height must agree with full import's auto height.
        sheet.Column(1).Width = 22; sheet.Column(3).Width = 18; sheet.Column(5).Width = 20;
        sheet.Column(6).Hide();
        for (int column = 1; column <= 7; column++) sheet.Cell(3, column).Value = "Header " + column;
        for (int row = 4; row <= 300; row++)
            for (int column = 1; column <= 7; column++) sheet.Cell(row, column).Value = row + column;
        var table = sheet.Range("A3:G300").CreateTable();
        table.Theme = XLTableTheme.TableStyleMedium2;
        table.ShowRowStripes = true;
        table.EmphasizeFirstColumn = true;
        MemoryStream source = new(); workbook.SaveAs(source);
        // Excel writers may omit font size; ClosedXML resolves that to its 11pt default.
        using (ZipArchive archive = new(source, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry("xl/styles.xml")!;
            XDocument styles;
            using (Stream stream = entry.Open()) styles = XDocument.Load(stream);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            foreach (var font in styles.Descendants(ns + "font").Where(font => font.Element(ns + "name")?.Attribute("val")?.Value == "Arial"))
                font.Element(ns + "sz")?.Remove();
            entry.Delete();
            using Stream output = archive.CreateEntry("xl/styles.xml").Open(); styles.Save(output);
        }
        source.Position = 0; return source;
    }
}
