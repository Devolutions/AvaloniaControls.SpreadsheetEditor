using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Themes.Fluent;
using Avalonia.VisualTree;

using ClosedXML.Excel;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

[assembly: AvaloniaTestApplication(typeof(SpreadsheetEditor.Tests.TestAppBuilder))]

namespace SpreadsheetEditor.Tests;

public sealed class TestApp : Application
{
    public override void Initialize() => this.Styles.Add(new FluentTheme());
}

public sealed class TestAppBuilder
{
    // Row-fit tests need the same font metrics and text shaping as the desktop renderer.
    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<TestApp>().UseSkia()
            .UseHeadless(new AvaloniaHeadlessPlatformOptions { UseHeadlessDrawing = false });
}

public sealed class WorkbookTests
{
    [AvaloniaFact]
    public void AddedSheetsAreNumberedAndNotifyTheHost()
    {
        EditorControl editor = new();
        int changes = 0;
        editor.ContentChanged += (_, _) => changes++;

        editor.AddSheet();
        editor.AddSheet();

        string prefix = EditorText.Get(EditorString.Sheet);
        Assert.Equal(new[] { $"{prefix}1", $"{prefix}2", $"{prefix}3" },
            editor.Sheets.Select(sheet => sheet.Name));
        Assert.Equal(2, changes);
    }

    [AvaloniaFact]
    public async Task SaveBeforeLazySheetsLoadPreservesAllOriginalData()
    {
        using MemoryStream source = CreateWorkbook();
        EditorControl editor = new();
        await editor.OpenXlsxAsync(source);

        Assert.Equal(3, editor.Sheets.Count);
        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        saved.Position = 0;
        var sheets = XlsxFile.ReadWorkbook(saved);

        Assert.Equal(new[] { "Sheet1", "Sheet2", "Sheet3" }, sheets.Select(sheet => sheet.Name));
        Assert.Equal("second sheet", sheets[1].Data.Cells[0][0].Text);
        Assert.Equal("third sheet", sheets[2].Data.Cells[0][0].Text);
        Assert.Equal(string.Empty, sheets[0].Data.Cells[0][0].Text);
    }

    [AvaloniaFact]
    public async Task EmptyFirstSheetFinishesLoadingAndEditsSaveAlongsideUnopenedSheets()
    {
        using MemoryStream source = CreateWorkbook();
        EditorControl editor = new();
        await editor.OpenXlsxAsync(source);

        var preview = editor.Sheets[0];
        Stopwatch timeout = Stopwatch.StartNew();
        while (ReferenceEquals(preview, editor.Sheets[0]) && timeout.Elapsed < TimeSpan.FromSeconds(10))
        {
            await Task.Delay(10);
        }
        Assert.NotSame(preview, editor.Sheets[0]);
        editor.ViewModel.SetCell(new CellAddress(0, 0), "edited blank sheet");

        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        saved.Position = 0;
        var sheets = XlsxFile.ReadWorkbook(saved);

        Assert.Equal("edited blank sheet", sheets[0].Data.Cells[0][0].Text);
        Assert.Equal("second sheet", sheets[1].Data.Cells[0][0].Text);
        Assert.Equal("third sheet", sheets[2].Data.Cells[0][0].Text);
    }

    [AvaloniaFact]
    public async Task NewSheetAddedWhileOriginalSheetsArePendingAppearsOnFirstReopen()
    {
        using MemoryStream source = CreateWorkbook();
        EditorControl editor = new();
        await editor.OpenXlsxAsync(source);
        int changes = 0;
        editor.ContentChanged += (_, _) => changes++;
        editor.AddSheet();
        editor.ViewModel.SetCell(new CellAddress(0, 0), "new sheet");

        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        saved.Position = 0;
        var sheets = XlsxFile.ReadWorkbook(saved);

        Assert.Equal(4, sheets.Count);
        Assert.Equal("Sheet4", sheets[3].Name);
        Assert.Equal("new sheet", sheets[3].Data.Cells[0][0].Text);
        Assert.Equal("second sheet", sheets[1].Data.Cells[0][0].Text);
        Assert.True(changes >= 1);
    }

    [AvaloniaFact]
    public async Task EditingSecondSheetAfterReopenPreservesEverySheetOnNextReopen()
    {
        using MemoryStream source = CreateWorkbook();
        EditorControl editor = new();
        await editor.OpenXlsxAsync(source);
        StackPanel tabs = editor.FindControl<StackPanel>("SheetTabs")!;
        ((Button)tabs.Children[1]).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Stopwatch timeout = Stopwatch.StartNew();
        while ((grid.ReadOnly || !ReferenceEquals(editor.ViewModel, editor.Sheets[1])) &&
               timeout.Elapsed < TimeSpan.FromSeconds(10))
        {
            await Task.Delay(10);
        }
        Assert.False(grid.ReadOnly);
        Assert.Same(editor.Sheets[1], editor.ViewModel);
        editor.ViewModel.SetCell(new CellAddress(0, 0), "edited second sheet");

        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        saved.Position = 0;
        EditorControl reopened = new();
        await reopened.OpenXlsxAsync(saved);

        Assert.Equal(3, reopened.Sheets.Count);
        using MemoryStream savedAgain = new();
        await reopened.SaveXlsxAsync(savedAgain);
        savedAgain.Position = 0;
        var sheets = XlsxFile.ReadWorkbook(savedAgain);
        Assert.Equal("edited second sheet", sheets[1].Data.Cells[0][0].Text);
        Assert.Equal("third sheet", sheets[2].Data.Cells[0][0].Text);
    }

    [AvaloniaFact]
    public void NewWorkbookEditsAndAllSheetsRoundTrip()
    {
        EditorControl editor = new();
        editor.NewSheet(1, 1);
        editor.ViewModel.SetCell(new CellAddress(0, 0), "first sheet");
        editor.AddSheet();
        editor.ViewModel.SetCell(new CellAddress(0, 0), "second sheet");
        editor.AddSheet();

        using MemoryStream saved = new();
        editor.SaveXlsx(saved);
        saved.Position = 0;
        var sheets = XlsxFile.ReadWorkbook(saved);

        Assert.Equal(3, sheets.Count);
        Assert.Equal("first sheet", sheets[0].Data.Cells[0][0].Text);
        Assert.Equal("second sheet", sheets[1].Data.Cells[0][0].Text);
        Assert.Equal(string.Empty, sheets[2].Data.Cells[0][0].Text);
    }

    [AvaloniaFact]
    public void ActiveCellEditMarksWorkbookDirtyOnceBeforeEditEnds()
    {
        EditorControl editor = new();
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        int changes = 0;
        editor.ContentChanged += (_, _) => changes++;

        grid.FocusCell();
        Assert.True(grid.IsEditing);
        editor.ViewModel.SetCell(new CellAddress(0, 0), "a");
        editor.ViewModel.SetCell(new CellAddress(0, 0), "ab");

        Assert.Equal(1, changes);
    }

    [AvaloniaFact]
    public void SavingDuringAnActiveEditDoesNotHideLaterChanges()
    {
        EditorControl editor = new();
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        int changes = 0;
        editor.ContentChanged += (_, _) => changes++;

        grid.FocusCell();
        editor.ViewModel.SetCell(new CellAddress(0, 0), "before save");
        Assert.Equal(1, changes);

        using MemoryStream firstSave = new();
        editor.SaveXlsx(firstSave);
        Assert.True(grid.IsEditing);
        editor.ViewModel.SetCell(new CellAddress(0, 0), "after save");
        Assert.Equal(2, changes);

        using MemoryStream secondSave = new();
        editor.SaveXlsx(secondSave);
        secondSave.Position = 0;
        Assert.Equal("after save", XlsxFile.ReadWorkbook(secondSave)[0].Data.Cells[0][0].Text);
    }

    [AvaloniaFact]
    public async Task WrappedHeaderAndRowHeightSurvivePreviewAndSave()
    {
        using XLWorkbook workbook = new();
        IXLWorksheet worksheet = workbook.Worksheets.Add("Headers");
        worksheet.Cell(2, 1).Value = "Long column heading";
        worksheet.Cell(2, 1).Style.Alignment.WrapText = true;
        worksheet.Row(2).Height = 28.5;
        using MemoryStream source = new();
        workbook.SaveAs(source);
        source.Position = 0;

        SheetData full = XlsxFile.Read(source);
        Assert.True(full.Cells[1][0].WrapText);
        Assert.InRange(full.RowHeights[1], 37.9, 38.1);

        source.Position = 0;
        EditorControl editor = new();
        await editor.OpenXlsxAsync(source);
        Assert.True(editor.ViewModel.Rows[1].Cells[0].WrapText);
        Assert.InRange(editor.ViewModel.Rows[1].Height!.Value, 37.9, 38.1);

        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        saved.Position = 0;
        using XLWorkbook reopened = new(saved);
        IXLWorksheet savedSheet = reopened.Worksheet("Headers");
        Assert.True(savedSheet.Cell(2, 1).Style.Alignment.WrapText);
        Assert.InRange(savedSheet.Row(2).Height, 28.4, 28.6);
    }

    [AvaloniaFact]
    public void WrappedCellsUseMultipleLinesWhenNoRowHeightIsSpecified()
    {
        const string heading = "Long wrapped heading with several words";
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[] { CellData.Plain(heading) with { WrapText = true } } },
            ColumnWidths = new[] { 80d },
        });
        Window window = new() { Content = editor, Width = 500, Height = 300 };
        window.Show();
        window.UpdateLayout();

        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        TextBlock text = Assert.Single(grid.GetVisualDescendants().OfType<TextBlock>(),
            block => block.Inlines?.OfType<Run>().Any(run => run.Text == heading) == true);
        Assert.Equal(TextWrapping.Wrap, text.TextWrapping);
        Assert.True(text.TextLayout.TextLines.Count > 1);
        Assert.True(text.TextLayout.TextLines.Sum(line => line.Length) >= heading.Length,
            "Every word must be included in the wrapped text layout.");
        Assert.True(text.Bounds.Height > 22);
        window.Close();
    }

    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public void WrappedHeaderDisplaysEveryWordWithinAnExcelRowHeight(bool hasFilter)
    {
        const string heading = "Avg session (min)";
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[] { CellData.Plain(heading) with
            {
                WrapText = true,
                Bold = true,
                FontFamily = "Tahoma",
                FontSize = 11 * 4.0 / 3.0,
                Alignment = TextAlignment.Center,
            } } },
            ColumnWidths = new[] { hasFilter ? 143d : 124d },
            RowHeights = new[] { 28 * 4.0 / 3.0 },
            AutoFilterRange = hasFilter ? new CellRange(0, 0, 1, 1) : null,
        });
        Window window = new() { Content = editor, Width = 500, Height = 300 };
        window.Show();
        window.UpdateLayout();

        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        TextBlock text = Assert.Single(grid.GetVisualDescendants().OfType<TextBlock>(),
            block => block.Inlines?.OfType<Run>().Any(run => run.Text == heading) == true);
        Assert.True(text.TextLayout.TextLines.Count > 1,
            $"Expected multiple lines, got {text.TextLayout.TextLines.Count} in {text.Bounds}.");
        Assert.True(text.TextLayout.TextLines.Sum(line => line.Length) >= heading.Length,
            "The wrapped header must include its last word.");
        window.Close();
    }

    [AvaloniaFact]
    public void EditedWrappedCellAndRowHeightAreWrittenToXlsx()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[] { CellData.Plain("Wrapped heading") with { WrapText = true } } },
            RowHeights = new[] { 38d },
        });

        using MemoryStream saved = new();
        editor.SaveXlsx(saved);
        saved.Position = 0;
        using XLWorkbook workbook = new(saved);
        Assert.True(workbook.Worksheet(1).Cell(1, 1).Style.Alignment.WrapText);
        Assert.InRange(workbook.Worksheet(1).Row(1).Height, 28.4, 28.6);
    }

    private static MemoryStream CreateWorkbook()
    {
        using XLWorkbook workbook = new();
        workbook.Worksheets.Add("Sheet1");
        workbook.Worksheets.Add("Sheet2").Cell(1, 1).Value = "second sheet";
        workbook.Worksheets.Add("Sheet3").Cell(1, 1).Value = "third sheet";
        MemoryStream stream = new();
        workbook.SaveAs(stream);
        stream.Position = 0;
        return stream;
    }
}
