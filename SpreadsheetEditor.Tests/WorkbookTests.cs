using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.Themes.Fluent;

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
    public static AppBuilder BuildAvaloniaApp() =>
        AppBuilder.Configure<TestApp>().UseHeadless(new AvaloniaHeadlessPlatformOptions());
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
