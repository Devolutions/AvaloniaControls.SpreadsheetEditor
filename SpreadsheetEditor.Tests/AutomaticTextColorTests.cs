namespace SpreadsheetEditor.Tests;

using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;

using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

public sealed class AutomaticTextColorTests
{
    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AutomaticResetsOnlyTheSelectedFontColorsAndRoundTripsAsAutomatic(bool dark)
    {
        EditorControl editor = new();
        CellData first = CellData.Plain("White text") with
        {
            Foreground = Colors.White, Background = Colors.Gray,
            Bold = true, Italic = true, Underline = true, NumberFormat = "@",
        };
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[]
            {
                first,
                CellData.Plain("Green text") with { Foreground = Colors.Green },
                CellData.Plain("Keep red") with { Foreground = Colors.Red },
            } },
        });
        Window window = new()
        {
            Content = editor, Width = 1100, Height = 500,
            RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light,
        };
        window.Show();
        try
        {
            await Frame(window);
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            CellData original = editor.ViewModel.Rows[0].Cells[0].ToData();
            grid.SelectCell(new CellAddress(0, 0));
            grid.FocusForSelection();
            window.KeyPress(Key.Right, RawInputModifiers.Shift, PhysicalKey.ArrowRight, null);
            Assert.Equal(new CellRange(0, 0, 1, 2), grid.SelectionRange);
            Button foreground = editor.FindControl<Button>("ForegroundButton")!;
            foreground.Flyout!.ShowAt(foreground);
            await Frame(window);
            Button automatic = editor.FindControl<Button>("AutomaticTextColorButton")!;
            Assert.Equal(EditorText.Get(EditorString.Automatic), automatic.GetVisualDescendants().OfType<TextBlock>().Single().Text);
            automatic.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Frame(window);

            Assert.Equal(original with { Foreground = null }, editor.ViewModel.Rows[0].Cells[0].ToData());
            Assert.Null(editor.ViewModel.Rows[0].Cells[1].ForegroundColor);
            Assert.Equal(Colors.Red, editor.ViewModel.Rows[0].Cells[2].ForegroundColor);
            TextBlock text = grid.GetVisualDescendants().OfType<TextBlock>()
                .Single(block => block.Name == "CellText" && Equals(block.Tag, new CellAddress(0, 0)));
            Assert.Equal(Color.FromRgb(0x20, 0x20, 0x20), Assert.IsAssignableFrom<ISolidColorBrush>(text.Foreground).Color);

            editor.Undo();
            Assert.Equal(Colors.White, editor.ViewModel.Rows[0].Cells[0].ForegroundColor);
            Assert.Equal(Colors.Green, editor.ViewModel.Rows[0].Cells[1].ForegroundColor);
            editor.Redo();
            Assert.Null(editor.ViewModel.Rows[0].Cells[0].ForegroundColor);
            Assert.Null(editor.ViewModel.Rows[0].Cells[1].ForegroundColor);

            using MemoryStream saved = new();
            await editor.SaveXlsxAsync(saved);
            saved.Position = 0;
            using (ZipArchive archive = new(saved, ZipArchiveMode.Read, leaveOpen: true))
            {
                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                using Stream stylesStream = archive.GetEntry("xl/styles.xml")!.Open();
                XDocument styles = XDocument.Load(stylesStream);
                using Stream sheetStream = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
                XDocument sheet = XDocument.Load(sheetStream);
                XElement cell = sheet.Descendants(ns + "c").Single(element => (string?)element.Attribute("r") == "A1");
                int styleId = (int?)cell.Attribute("s") ?? 0;
                XElement xf = styles.Root!.Element(ns + "cellXfs")!.Elements().ElementAt(styleId);
                XElement font = styles.Root.Element(ns + "fonts")!.Elements().ElementAt((int)xf.Attribute("fontId")!);
                Assert.Equal("1", (string?)font.Element(ns + "color")?.Attribute("auto"));
            }
            saved.Position = 0;
            CellData[] cells = XlsxFile.ReadWorkbook(saved)[0].Data.Cells[0].Take(3).ToArray();
            Assert.Null(cells[0].Foreground);
            Assert.Null(cells[1].Foreground);
            Assert.Equal(Colors.Red, cells[2].Foreground);
            Assert.Equal(Colors.Gray, cells[0].Background);
            Assert.True(cells[0].Bold && cells[0].Italic && cells[0].Underline);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        }
    }

    [AvaloniaFact]
    public void ReadOnlyRejectsAutomaticEvenIfTheFlyoutWasAlreadyOpen()
    {
        EditorControl editor = new();
        editor.ViewModel.SetCell(new CellAddress(0, 0), "White text");
        editor.ViewModel.Rows[0].Cells[0].SetWholeForeground(Colors.White);
        editor.IsReadOnly = true;
        Button automatic = editor.FindControl<Button>("AutomaticTextColorButton")!;
        Assert.False(automatic.IsEnabled);
        automatic.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        Assert.Equal(Colors.White, editor.ViewModel.Rows[0].Cells[0].ForegroundColor);
        editor.NewSheet(1, 1);
    }

    [AvaloniaFact]
    public async Task AutomaticFontDoesNotExpandTheWorksheetWhenSavedAgain()
    {
        EditorControl editor = new();
        editor.NewSheet(1000, 26);
        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        using ZipArchive archive = new(saved, ZipArchiveMode.Read);
        using Stream sheet = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XDocument xml = XDocument.Load(sheet);
        XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        int originalCount = xml.Descendants(ns + "c").Count();
        saved.Position = 0;
        await editor.OpenXlsxAsync(saved);
        using MemoryStream resaved = new();
        await editor.SaveXlsxAsync(resaved);
        using ZipArchive secondArchive = new(resaved, ZipArchiveMode.Read);
        using Stream secondSheet = secondArchive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        Assert.InRange(XDocument.Load(secondSheet).Descendants(ns + "c").Count(), 0, originalCount);
        editor.NewSheet(1, 1);
    }

    [AvaloniaFact]
    public async Task AutomaticOverridesTableFontColorOnImport()
    {
        SheetData data = new()
        {
            Cells = new[]
            {
                new[] { CellData.Plain("Automatic"), CellData.Plain("White") with { Foreground = Colors.White } },
                new[] { CellData.Plain("First"), CellData.Plain("Second") },
            },
            Tables = new[] { new SheetTable { Name = "Test", Range = new CellRange(0, 0, 2, 2), StyleName = "TableStyleMedium4" } },
        };
        EditorControl editor = new();
        editor.LoadSheet(data);
        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        saved.Position = 0;
        SheetData preview = XlsxFile.ReadWorkbook(saved)[0].Data;
        Assert.Null(preview.Cells[0][0].Foreground);
        Assert.Equal(Colors.White, preview.Cells[0][1].Foreground);
        saved.Position = 0;
        using ClosedXML.Excel.XLWorkbook workbook = new(saved);
        saved.Position = 0;
        await editor.OpenXlsxAsync(saved);
        Assert.Null(editor.ViewModel.Rows[0].Cells[0].ForegroundColor);
        Assert.Equal(Colors.White, editor.ViewModel.Rows[0].Cells[1].ForegroundColor);
        editor.NewSheet(1, 1);
    }

    [AvaloniaFact]
    public void AutomaticClearsRichRunColorsAndPreservesTheirOtherFormatting()
    {
        EditorControl editor = new();
        editor.NewSheet(2, 2);
        editor.ViewModel.Rows[0].Cells[0].SetRichRuns(new()
        {
            new CellRun("White", true, false, false, Colors.White, null),
            new CellRun(" red", false, true, false, Colors.Red, null),
        });
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        grid.SelectCell(new CellAddress(0, 0));
        grid.FocusForSelection();
        editor.FindControl<Button>("AutomaticTextColorButton")!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        Assert.All(editor.ViewModel.Rows[0].Cells[0].EditRuns, run => Assert.Null(run.Foreground));
        Assert.True(editor.ViewModel.Rows[0].Cells[0].EditRuns[0].Bold);
        Assert.True(editor.ViewModel.Rows[0].Cells[0].EditRuns[1].Italic);
        editor.NewSheet(1, 1);
    }

    private static async Task Frame(Window window)
    {
        window.UpdateLayout();
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Render);
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
    }
}
