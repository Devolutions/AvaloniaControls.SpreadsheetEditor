using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;

using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Formulas;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class TableFilterTests
{
    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public void WorksheetFiltersAndBlankTableFiltersPersistTheirCriteria(bool worksheetFilter)
    {
        SheetData original = CreateSheet();
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = original.Cells,
            Tables = worksheetFilter ? Array.Empty<SheetTable>() : original.Tables,
            AutoFilterRange = worksheetFilter ? new CellRange(2, 1, 5, 3) : null,
        });
        editor.ViewModel.FilterColumn(2, string.Empty, 2);
        using MemoryStream saved = new();
        editor.SaveXlsx(saved);
        saved.Position = 0;
        SheetData reopened = XlsxFile.Read(saved);
        Assert.Equal(worksheetFilter ? 0 : 3, reopened.Tables.Count);
        editor.LoadSheet(reopened);
        Assert.True(editor.ViewModel.HasColumnFilter(2, 2));
        Assert.True(editor.ViewModel.Rows[3].IsHidden);
        Assert.False(editor.ViewModel.Rows[6].IsHidden);
        editor.ViewModel.ClearTableFilters(2, 2);
        Assert.False(editor.ViewModel.Rows[3].IsHidden);
    }

    [AvaloniaFact]
    public void SavingChartsAndTablesTogetherKeepsValidWorksheetOrdering()
    {
        SheetData source = CreateSheet();
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = source.Cells, Tables = source.Tables,
            Charts = new[] { new SheetChart
            {
                Kind = ChartKind.Line,
                Series = new[] { new ChartSeriesSpec("Values", new CellRange(3, 3, 4, 1)) },
                AnchorRow = 1, AnchorColumn = 8,
            } },
        });
        using MemoryStream saved = new();
        editor.SaveXlsx(saved);
        saved.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(saved, false))
        {
            var errors = new OpenXmlValidator().Validate(document)
                .Where(error => error.Part?.Uri.OriginalString.Contains("/worksheets/", StringComparison.Ordinal) == true ||
                                error.Part?.Uri.OriginalString.Contains("/tables/", StringComparison.Ordinal) == true).ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
        }
        saved.Position = 0;
        SheetData reopened = XlsxFile.Read(saved);
        Assert.Equal(3, reopened.Tables.Count);
        Assert.Single(reopened.Charts);
    }

    [AvaloniaFact]
    public async Task HeaderFilterMenuUsesItsOwnTablesValues()
    {
        EditorControl editor = new();
        editor.LoadSheet(CreateSheet());
        Window window = new() { Content = editor, Width = 1100, Height = 800 };
        window.Show();
        try
        {
            await Frame(window);
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Border button = grid.GetVisualDescendants().OfType<Border>().Single(border =>
                border.Name == "ColumnFilterButton" && Equals(border.Tag, new CellAddress(12, 2)));
            MenuFlyout menu = Assert.IsType<MenuFlyout>(button.ContextFlyout);
            menu.ShowAt(button);
            MenuItem filter = menu.Items.OfType<MenuItem>().ElementAt(2);
            filter.IsSubMenuOpen = true;
            await Frame(window);
            Assert.DoesNotContain(filter.Items.OfType<MenuItem>(), item => Equals(item.Header, "Charlie"));
            MenuItem sales = filter.Items.OfType<MenuItem>().Single(item => Equals(item.Header, "Sales"));
            sales.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            menu.Hide();
            Assert.True(editor.ViewModel.Rows[13].IsHidden);
            Assert.False(editor.ViewModel.Rows[14].IsHidden);
            Assert.False(editor.ViewModel.Rows[4].IsHidden);
        }
        finally { editor.NewSheet(1, 1); window.Close(); await Frame(window); }
    }

    [AvaloniaFact]
    public void NativeExcelTableDefinitionsAndFiltersRoundTrip()
    {
        EditorControl editor = new();
        editor.LoadSheet(CreateSheet());
        editor.ViewModel.FilterColumn(2, "Ops", 2);
        editor.ViewModel.FilterColumn(2, "Sales", 12);
        using MemoryStream saved = new();
        editor.SaveXlsx(saved);
        saved.Position = 0;
        using (XLWorkbook workbook = new(saved))
        {
            Assert.Equal(3, workbook.Worksheet(1).Tables.Count());
            Assert.Equal("B3:D7", workbook.Worksheet(1).Table("People").RangeAddress.ToStringRelative());
            Assert.True(workbook.Worksheet(1).Table("People").ShowAutoFilter);
            Assert.False(workbook.Worksheet(1).AutoFilter.IsEnabled);
        }
        saved.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(saved, false))
        {
            var errors = new OpenXmlValidator().Validate(document).ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(error => error.Path?.XPath + ": " + error.Description)));
        }
        saved.Position = 0;
        SheetData reopened = XlsxFile.Read(saved);
        Assert.Null(reopened.AutoFilterRange);
        Assert.Equal(3, reopened.Tables.Count);
        Assert.Equal(new CellRange(2, 1, 5, 3), reopened.Tables[0].AutoFilterRange);
        Assert.Equal("TableStyleMedium4", reopened.Tables[0].StyleName);
        editor.LoadSheet(reopened);
        Assert.True(editor.ViewModel.HasColumnFilter(2, 2));
        Assert.True(editor.ViewModel.HasColumnFilter(2, 12));
        Assert.True(editor.ViewModel.Rows[4].IsHidden);
        Assert.True(editor.ViewModel.Rows[13].IsHidden);
        editor.ViewModel.ClearTableFilters(2, 2);
        Assert.False(editor.ViewModel.Rows[4].IsHidden);
        Assert.True(editor.ViewModel.Rows[13].IsHidden);
        editor.ViewModel.ClearTableFilters(2, 12);
        Assert.False(editor.ViewModel.Rows[13].IsHidden);
    }

    [AvaloniaFact]
    public void SortingMovesOnlyTableCellsAndKeepsTheirFormattingAndFormulas()
    {
        SheetData source = CreateSheet();
        var rows = source.Cells.Select(row => row.ToArray()).ToArray();
        for (int row = 3; row <= 6; row++)
            rows[row][3] = CellData.Plain($"=LEN(B{row + 1})") with
            { FormulaResult = rows[row][1].Text.Length.ToString(), FormulaNumericValue = rows[row][1].Text.Length };
        rows[4][1] = rows[4][1] with { Bold = true, Comment = "Alice comment", Hyperlink = "https://example.com/alice" };
        EditorControl editor = new();
        editor.LoadSheet(CopyWithCells(source, rows));
        editor.ViewModel.Rows[4].Cells[1].SetRichRuns(new List<CellRun>
        {
            new("Al", true, false, false, Colors.Red, null), new("ice", false, true, false, null, null),
        });
        editor.ViewModel.SortRows(1, 2, false);
        Assert.Equal(new[] { "Alice", "Bob", "Charlie", "Diana" }, Enumerable.Range(3, 4).Select(row => editor.ViewModel.Rows[row].Cells[1].Text));
        Assert.True(editor.ViewModel.Rows[3].Cells[1].Bold);
        Assert.Equal("Alice comment", editor.ViewModel.Rows[3].Cells[1].Comment);
        Assert.Equal("https://example.com/alice", editor.ViewModel.Rows[3].Cells[1].Hyperlink);
        Assert.Equal("=LEN(B4)", editor.ViewModel.Rows[3].Cells[3].Text);
        Assert.Equal(5, editor.ViewModel.Rows[3].Cells[3].CachedNumericValue);
        Assert.Equal(2, editor.ViewModel.Rows[3].Cells[1].Runs!.Count);
        Assert.Equal(Colors.Red, editor.ViewModel.Rows[3].Cells[1].Runs![0].Foreground);
        for (int row = 0; row < source.Cells.Count; row++)
        for (int column = 0; column < source.Cells[row].Count; column++)
            if (row < 3 || row > 6 || column < 1 || column > 3)
                Assert.Equal(rows[row][column].Text, editor.ViewModel.Rows[row].Cells[column].Text);
    }

    [AvaloniaFact]
    public void MultipleTablesFilterIndependentlyAndKeepManualHiddenRows()
    {
        EditorControl editor = new();
        editor.LoadSheet(CreateSheet());
        editor.ViewModel.SetHiddenRows(new[] { 8 });
        Assert.Equal(new[] { "", "Ops", "Sales" }, editor.ViewModel.ColumnValues(2, 2));
        editor.ViewModel.FilterColumn(2, "Ops", 2);
        Assert.False(editor.ViewModel.Rows[3].IsHidden);
        Assert.True(editor.ViewModel.Rows[4].IsHidden);
        Assert.False(editor.ViewModel.Rows[13].IsHidden);
        editor.ViewModel.FilterColumn(2, "Sales", 12);
        Assert.True(editor.ViewModel.Rows[13].IsHidden);
        Assert.False(editor.ViewModel.Rows[14].IsHidden);
        editor.ViewModel.ClearTableFilters(2, 2);
        Assert.False(editor.ViewModel.Rows[4].IsHidden);
        Assert.True(editor.ViewModel.Rows[13].IsHidden);
        Assert.True(editor.ViewModel.Rows[8].IsHidden);
        editor.ViewModel.FilterColumn(2, null, 12);
        Assert.False(editor.ViewModel.Rows[13].IsHidden);
        Assert.True(editor.ViewModel.Rows[8].IsHidden);
    }

    [AvaloniaFact]
    public void TablesBesideEachOtherKeepTheirOwnColumnFilters()
    {
        EditorControl editor = new();
        editor.LoadSheet(CreateSheet());
        editor.ViewModel.FilterColumn(2, "Ops", 2);
        editor.ViewModel.FilterColumn(6, "Active", 2);
        editor.ViewModel.ClearTableFilters(2, 2);
        Assert.False(editor.ViewModel.HasColumnFilter(2, 2));
        Assert.True(editor.ViewModel.HasColumnFilter(6, 2));
        Assert.True(editor.ViewModel.Rows[5].IsHidden);
        Assert.Equal(new[] { "Active", "Inactive" }, editor.ViewModel.ColumnValues(6, 2));
    }

    [AvaloniaFact]
    public void BlankFilteringAndNumericSortingStayInsideTheTable()
    {
        EditorControl editor = new();
        editor.LoadSheet(CreateSheet());
        editor.ViewModel.FilterColumn(2, string.Empty, 2);
        Assert.True(editor.ViewModel.Rows[3].IsHidden);
        Assert.False(editor.ViewModel.Rows[6].IsHidden);
        editor.ViewModel.FilterColumn(2, null, 2);
        editor.ViewModel.SortRows(3, 2, true);
        Assert.Equal(new[] { "40", "30", "20", "10" }, Enumerable.Range(3, 4).Select(row => editor.ViewModel.Rows[row].Cells[3].Text));
        Assert.Equal("Below table", editor.ViewModel.Rows[8].Cells[1].Text);
        Assert.Equal("Zebra", editor.ViewModel.Rows[13].Cells[1].Text);
    }

    [AvaloniaFact]
    public void RowInsertionAndDeletionMoveTableRangesAndFilterHeadersTogether()
    {
        EditorControl editor = new();
        editor.LoadSheet(CreateSheet());
        editor.ViewModel.FilterColumn(2, "Ops", 2);
        editor.ViewModel.InsertRows(1, 2);
        Assert.Equal(new CellRange(4, 1, 5, 3), editor.ViewModel.Tables[0].AutoFilterRange);
        Assert.True(editor.ViewModel.HasColumnFilter(2, 4));
        Assert.True(editor.ViewModel.Rows[6].IsHidden);
        editor.ViewModel.DeleteRows(1, 2);
        Assert.Equal(new CellRange(2, 1, 5, 3), editor.ViewModel.Tables[0].AutoFilterRange);
        Assert.True(editor.ViewModel.HasColumnFilter(2, 2));
        Assert.True(editor.ViewModel.Rows[4].IsHidden);
    }

    [AvaloniaFact]
    public void HeaderlessTablesAndTotalsDoNotShiftOrOverwriteSurroundingCellsOnSave()
    {
        SheetData original = CreateSheet();
        var rows = original.Cells.Select(row => row.ToArray()).ToArray();
        rows[7][1] = CellData.Plain("Total");
        rows[7][3] = CellData.Plain("100");
        SheetTable[] tables = original.Tables.Select(table => table.Name == "People" ? table with
            { Range = new CellRange(2, 1, 6, 3), ShowTotalsRow = true } : table.Name == "States" ? table with
            { Range = new CellRange(3, 5, 4, 2), ShowHeaderRow = false, AutoFilterRange = null, ColumnNames = new[] { "Key", "State" } } : table).ToArray();
        EditorControl editor = new();
        editor.LoadSheet(new SheetData { Cells = rows, Tables = tables });
        using MemoryStream saved = new();
        editor.SaveXlsx(saved);
        saved.Position = 0;
        SheetData reopened = XlsxFile.Read(saved);
        Assert.True(reopened.Tables.Single(table => table.Name == "People").ShowTotalsRow);
        Assert.False(reopened.Tables.Single(table => table.Name == "States").ShowHeaderRow);
        Assert.Equal(new CellRange(2, 1, 5, 3), reopened.Tables[0].AutoFilterRange);
        for (int row = 0; row < rows.Length; row++)
        for (int column = 0; column < rows[row].Length; column++)
            Assert.Equal(rows[row][column].Text, row < reopened.Cells.Count && column < reopened.Cells[row].Count ? reopened.Cells[row][column].Text : string.Empty);
        editor.ViewModel.SortRows(3, 2, false);
        Assert.Equal("Total", editor.ViewModel.Rows[7].Cells[1].Text);
        Assert.Equal("100", editor.ViewModel.Rows[7].Cells[3].Text);
    }

    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HeaderDropdownOpensTheCorrectTablesSortMenuAndRespectsReadOnly(bool readOnly)
    {
        EditorControl editor = new();
        editor.LoadSheet(CreateSheet());
        Window window = new() { Content = editor, Width = 1100, Height = 800 };
        window.Show();
        try
        {
            await Frame(window);
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Border[] buttons = grid.GetVisualDescendants().OfType<Border>().Where(border => border.Name == "ColumnFilterButton").ToArray();
            Assert.Equal(8, buttons.Length);
            Border button = buttons.Single(button => Equals(button.Tag, new CellAddress(12, 1)));
            MenuFlyout menu = Assert.IsType<MenuFlyout>(button.ContextFlyout);
            editor.IsReadOnly = readOnly;
            Point point = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window)!.Value;
            window.MouseDown(point, MouseButton.Left);
            window.MouseUp(point, MouseButton.Left);
            await Frame(window);
            Assert.True(menu.IsOpen);
            MenuItem ascending = menu.Items.OfType<MenuItem>().First();
            Assert.Equal(!readOnly, ascending.IsEnabled);
            ascending.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            menu.Hide();
            Assert.Equal(readOnly ? "Zebra" : "Wolf", editor.ViewModel.Rows[13].Cells[1].Text);
            Assert.Equal("Charlie", editor.ViewModel.Rows[3].Cells[1].Text);
        }
        finally { editor.NewSheet(1, 1); window.Close(); await Frame(window); }
    }

    [AvaloniaFact]
    public async Task LazyOpenAndSaveBeforeFullLoadKeepTableHeaderMetadata()
    {
        EditorControl source = new();
        source.LoadSheet(CreateSheet());
        using MemoryStream input = new();
        source.SaveXlsx(input);
        input.Position = 0;
        EditorControl editor = new();
        await editor.OpenXlsxAsync(input);
        Assert.Equal(3, editor.ViewModel.Tables.Count); // available in the first-screen preview too
        Assert.Equal(new CellRange(2, 1, 5, 3), editor.ViewModel.FilterHeaderRange(2, 2));
        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        saved.Position = 0;
        Assert.Equal(3, XlsxFile.Read(saved).Tables.Count);
        Stopwatch timeout = Stopwatch.StartNew();
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        while (grid.ReadOnly && timeout.Elapsed < TimeSpan.FromSeconds(10)) await Task.Delay(10);
        Assert.False(grid.ReadOnly);
        Assert.Equal(3, editor.ViewModel.Tables.Count);
        editor.NewSheet(1, 1);
        await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
    }

    private static SheetData CopyWithCells(SheetData original, CellData[][] cells) => new()
    { Cells = cells, Tables = original.Tables, ColumnWidths = original.ColumnWidths };

    private static SheetData CreateSheet()
    {
        CellData[][] cells = Enumerable.Range(0, 22).Select(_ => Enumerable.Repeat(CellData.Plain(string.Empty), 11).ToArray()).ToArray();
        cells[0][0] = CellData.Plain("Introduction, not a table header");
        cells[8][1] = CellData.Plain("Below table");
        string[] names = { "Charlie", "Alice", "Bob", "Diana" };
        string[] departments = { "Ops", "Sales", "Ops", "" };
        string[] bottomNames = { "Zebra", "Xeno", "Yolo", "Wolf" };
        int[] values = { 30, 10, 20, 40 };
        for (int offset = 0; offset < 3; offset++)
        {
            cells[2][1 + offset] = CellData.Plain(new[] { "Name", "Department", "Value" }[offset]);
            cells[12][1 + offset] = cells[2][1 + offset];
        }
        cells[2][5] = CellData.Plain("Key"); cells[2][6] = CellData.Plain("State");
        for (int index = 0; index < 4; index++)
        {
            cells[3 + index][1] = CellData.Plain(names[index]);
            cells[3 + index][2] = CellData.Plain(departments[index]);
            cells[3 + index][3] = CellData.Plain(values[index].ToString());
            cells[3 + index][5] = CellData.Plain("ID-" + index);
            cells[3 + index][6] = CellData.Plain(index == 2 ? "Inactive" : "Active");
            cells[3 + index][8] = CellData.Plain("Unrelated note " + index);
            cells[13 + index][1] = CellData.Plain(bottomNames[index]);
            cells[13 + index][2] = CellData.Plain(index == 1 ? "Sales" : "Ops");
            cells[13 + index][3] = CellData.Plain((100 * (index + 1)).ToString());
        }
        return new SheetData
        {
            Cells = cells,
            ColumnWidths = Enumerable.Repeat(85d, 11).ToArray(),
            Tables = new[]
            {
                new SheetTable { Name = "People", Range = new CellRange(2, 1, 5, 3), AutoFilterRange = new CellRange(2, 1, 5, 3),
                    StyleName = "TableStyleMedium4", ShowRowStripes = true },
                new SheetTable { Name = "States", Range = new CellRange(2, 5, 5, 2), AutoFilterRange = new CellRange(2, 5, 5, 2) },
                new SheetTable { Name = "MorePeople", Range = new CellRange(12, 1, 5, 3), AutoFilterRange = new CellRange(12, 1, 5, 3) },
            },
        };
    }

    private static async Task Frame(Window window)
    {
        await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
        window.UpdateLayout();
    }
}
