using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.ViewModels;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class FillHandleTests
{
    [Theory]
    [InlineData("=A1+$A1+A$1+$A$1", "=C3+$A3+C$1+$A$1")]
    [InlineData("=SUM(A1:B2)", "=SUM(C3:D4)")]
    [InlineData("='Data A1'!A1+Data!$B2", "='Data A1'!C3+Data!$B4")]
    [InlineData("=IF(A1=\"B2\",\"A1 \"\"B2\"\"\",LOG10(C2))", "=IF(C3=\"B2\",\"A1 \"\"B2\"\"\",LOG10(E4))")]
    [InlineData("=SUM(A:A,$B:$B,1:2,$3:$3)", "=SUM(C:C,$B:$B,3:4,$3:$3)")]
    [InlineData("=Table1[Column1]+A1", "=Table1[Column1]+C3")]
    public void FillingCopiesFormulasWithRelativeAbsoluteAndMixedReferences(string formula, string expected)
    {
        SpreadsheetViewModel sheet = new(6, 6);
        sheet.SetCell(new CellAddress(3, 3), formula);
        Assert.True(sheet.FillCells(new CellRange(3, 3, 1, 1), new CellRange(3, 3, 3, 3)));
        Assert.Equal(formula, sheet.Rows[3].Cells[3].Text);
        Assert.Equal(expected, sheet.Rows[5].Cells[5].Text);
    }

    [Fact]
    public void FillingUpProducesReferenceErrorsRatherThanWrappingInvalidReferences()
    {
        SpreadsheetViewModel sheet = new(2, 2);
        sheet.SetCell(new CellAddress(1, 1), "=A1");
        Assert.True(sheet.FillCells(new CellRange(1, 1, 1, 1), new CellRange(0, 1, 2, 1)));
        Assert.Equal("=#REF!", sheet.Rows[0].Cells[1].Text);
    }

    [Fact]
    public void FillingRepeatsAWholeBlockWithFormattingCommentsLinksAndOneUndoPoint()
    {
        SpreadsheetViewModel sheet = new(7, 2);
        CellData first = CellData.Plain("One") with
        {
            Bold = true, Background = Colors.Yellow, Foreground = Colors.Red,
            NumberFormat = "0.0%", FontSize = 14, WrapText = true, Comment = "Note", Hyperlink = "https://example.com",
        };
        sheet.Rows[0].Cells[0].ApplyData(first);
        sheet.SetCell(new CellAddress(0, 1), "Two");
        sheet.SetCell(new CellAddress(1, 0), "Three");
        sheet.SetCell(new CellAddress(1, 1), "Four");
        sheet.SetCell(new CellAddress(2, 0), "Previous value");
        Assert.True(sheet.FillCells(new CellRange(0, 0, 2, 2), new CellRange(0, 0, 7, 2)));
        Assert.Equal(first, sheet.Rows[6].Cells[0].ToData());
        Assert.Equal("Four", sheet.Rows[5].Cells[1].Text);
        sheet.Undo();
        Assert.Equal("Previous value", sheet.Rows[2].Cells[0].Text);
        Assert.Equal(string.Empty, sheet.Rows[6].Cells[0].Text);
        Assert.False(sheet.CanUndo);
        sheet.Redo();
        Assert.Equal(first, sheet.Rows[6].Cells[0].ToData());
    }

    [Fact]
    public void FillingLeftRepeatsTheBlockInItsOriginalOrder()
    {
        SpreadsheetViewModel sheet = new(1, 6);
        sheet.SetCell(new CellAddress(0, 4), "One");
        sheet.SetCell(new CellAddress(0, 5), "Two");
        sheet.FillCells(new CellRange(0, 4, 1, 2), new CellRange(0, 0, 1, 6));
        Assert.Equal(new[] { "One", "Two", "One", "Two", "One", "Two" }, sheet.Rows[0].Cells.Select(cell => cell.Text));
    }

    [Fact]
    public void FillingIntoMergedCellsIsRejectedWithoutChangingTheSheet()
    {
        SpreadsheetViewModel sheet = new(3, 2);
        sheet.SetCell(new CellAddress(0, 0), "Keep");
        sheet.SetMerges(new[] { new MergeRange(1, 0, 1, 2) });
        Assert.False(sheet.FillCells(new CellRange(0, 0, 1, 1), new CellRange(0, 0, 3, 1)));
        Assert.Equal(string.Empty, sheet.Rows[1].Cells[0].Text);
        Assert.False(sheet.CanUndo);
    }

    [AvaloniaTheory]
    [InlineData(3, 0)]
    [InlineData(-2, 0)]
    [InlineData(0, 3)]
    [InlineData(0, -2)]
    public async Task DraggingTheHandleCopiesOnReleaseInEveryDirection(int rows, int columns)
    {
        EditorControl editor = CreateEditor();
        editor.ViewModel.SetCell(new CellAddress(2, 2), "Copy me");
        Window window = new() { Content = editor, Width = 800, Height = 500 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            grid.SelectCell(new CellAddress(2, 2));
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
            window.UpdateLayout();
            Border handle = Handle(grid);
            Assert.True(handle.IsVisible);
            Point start = handle.TranslatePoint(new Point(5, 5), window)!.Value;
            Point end = new(start.X + columns * 80, start.Y + rows * 24);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            Assert.Equal(string.Empty, editor.ViewModel.Rows[2 + rows].Cells[2 + columns].Text);
            Assert.True(grid.GetVisualDescendants().OfType<Border>().Single(border => border.Name == "FillPreview").IsVisible);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal("Copy me", editor.ViewModel.Rows[2 + rows].Cells[2 + columns].Text);
            Assert.Equal(new CellRange(Math.Min(2, 2 + rows), Math.Min(2, 2 + columns), Math.Abs(rows) + 1, Math.Abs(columns) + 1),
                grid.SelectionRange);
            editor.Undo();
            Assert.Equal(string.Empty, editor.ViewModel.Rows[2 + rows].Cells[2 + columns].Text);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaFact]
    public void EscapeCancelsAndReadOnlyOrEditingHidesTheHandle()
    {
        EditorControl editor = CreateEditor();
        editor.ViewModel.SetCell(new CellAddress(0, 0), "Keep");
        Window window = new() { Content = editor, Width = 600, Height = 400 };
        window.Show();
        try
        {
            window.UpdateLayout();
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Point start = Handle(grid).TranslatePoint(new Point(5, 5), window)!.Value;
            Point end = new(start.X, start.Y + 72);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.Escape, null);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal(string.Empty, editor.ViewModel.Rows[3].Cells[0].Text);
            Assert.False(editor.ViewModel.CanUndo);
            editor.IsReadOnly = true;
            Assert.False(Handle(grid).IsVisible);
            editor.IsReadOnly = false;
            Assert.True(Handle(grid).IsVisible);
            grid.FocusCell();
            Assert.False(Handle(grid).IsVisible);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaFact]
    public void ClickingTheHandleWithoutDraggingDoesNotChangeAnything()
    {
        EditorControl editor = CreateEditor();
        Window window = new() { Content = editor, Width = 600, Height = 400 };
        window.Show();
        try
        {
            window.UpdateLayout();
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Point start = Handle(grid).TranslatePoint(new Point(6, 6), window)!.Value;
            window.MouseDown(start, MouseButton.Left);
            window.MouseUp(start, MouseButton.Left);
            Assert.Equal(new CellRange(0, 0, 1, 1), grid.SelectionRange);
            Assert.False(editor.ViewModel.CanUndo);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaTheory]
    [InlineData(0.5)]
    [InlineData(2.0)]
    public async Task DraggingAMulticellSelectionRepeatsItsPatternAtDifferentZoomLevels(double zoom)
    {
        EditorControl editor = CreateEditor();
        editor.ViewModel.SetCell(new CellAddress(0, 0), "A");
        editor.ViewModel.SetCell(new CellAddress(0, 1), "B");
        editor.ViewModel.SetCell(new CellAddress(1, 0), "C");
        editor.ViewModel.SetCell(new CellAddress(1, 1), "D");
        Window window = new() { Content = editor, Width = 800, Height = 600 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            grid.Zoom = zoom;
            grid.FocusForSelection();
            window.KeyPress(Key.Right, RawInputModifiers.Shift, PhysicalKey.ArrowRight, null);
            window.KeyPress(Key.Down, RawInputModifiers.Shift, PhysicalKey.ArrowDown, null);
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
            window.UpdateLayout();
            Assert.Equal(new CellRange(0, 0, 2, 2), grid.SelectionRange);
            Point start = Handle(grid).TranslatePoint(new Point(5, 5), window)!.Value;
            Point end = new(start.X, start.Y + 4 * 24 * zoom);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            window.MouseUp(end, MouseButton.Left);
            Assert.Equal("A", editor.ViewModel.Rows[4].Cells[0].Text);
            Assert.Equal("D", editor.ViewModel.Rows[5].Cells[1].Text);
            Assert.Equal(new CellRange(0, 0, 6, 2), grid.SelectionRange);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DraggingPastTheViewportScrollsAndKeepsPointerCaptureDuringRendering(bool horizontal)
    {
        EditorControl editor = CreateEditor();
        editor.ViewModel.SetCell(new CellAddress(0, 0), "Copy me");
        Window window = new() { Content = editor, Width = 600, Height = 400 };
        window.Show();
        try
        {
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
            window.UpdateLayout();
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Border handle = Handle(grid);
            ScrollViewer scroller = handle.GetVisualAncestors().OfType<ScrollViewer>().First();
            Point start = handle.TranslatePoint(new Point(5, 5), window)!.Value;
            Point end = horizontal ? new Point(window.Bounds.Width + 20, start.Y) : new Point(start.X, window.Bounds.Height + 20);
            window.MouseDown(start, MouseButton.Left);
            window.MouseMove(end);
            Stopwatch timeout = Stopwatch.StartNew();
            while ((horizontal ? scroller.Offset.X : scroller.Offset.Y) < 100 && timeout.Elapsed < TimeSpan.FromSeconds(5))
            {
                await Task.Delay(30);
                window.UpdateLayout();
            }
            Assert.True((horizontal ? scroller.Offset.X : scroller.Offset.Y) >= 100);
            Assert.Contains(scroller, Handle(grid).GetVisualAncestors());
            window.MouseUp(end, MouseButton.Left);
            Assert.True(horizontal ? grid.SelectionRange.ColumnSpan > 7 : grid.SelectionRange.RowSpan > 8);
            Assert.Equal("Copy me", editor.ViewModel.Rows[grid.SelectionRange.RowSpan - 1].Cells[grid.SelectionRange.ColumnSpan - 1].Text);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    [AvaloniaFact]
    public async Task FillingAnInputRecalculatesOtherSheetsAndNotifiesTheHostOnce()
    {
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Data", new SheetData { Cells = new[] { new[] { CellData.Plain("1") }, new[] { CellData.Plain("2") } } }),
            ("Summary", new SheetData { Cells = new[] { new[] { CellData.Plain("=SUM(Data!A1:A5)") with
                { FormulaResult = "3", FormulaNumericValue = 3 } } } }),
        });
        int notifications = 0;
        editor.ContentChanged += (_, _) => notifications++;
        try
        {
            editor.ViewModel.FillCells(new CellRange(0, 0, 1, 1), new CellRange(0, 0, 5, 1));
            Assert.Equal(1, notifications);
            Stopwatch timeout = Stopwatch.StartNew();
            while (editor.Sheets[1].Rows[0].Cells[0].CachedNumericValue != 5 && timeout.Elapsed < TimeSpan.FromSeconds(5))
            {
                await Task.Delay(5);
            }
            Assert.Equal(5, editor.Sheets[1].Rows[0].Cells[0].CachedNumericValue);
        }
        finally
        {
            editor.NewSheet(1, 1);
        }
    }

    [AvaloniaFact]
    public async Task CopiedFormulasRecalculateAndKeepTheirCachedResultsWhenSaved()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[]
            {
                new[] { CellData.Plain("1"), CellData.Plain("=A1*2") with { FormulaResult = "2", FormulaNumericValue = 2 } },
                new[] { CellData.Plain("3"), CellData.Plain(string.Empty) },
                new[] { CellData.Plain("4"), CellData.Plain(string.Empty) },
            },
        });
        try
        {
            editor.ViewModel.FillCells(new CellRange(0, 1, 1, 1), new CellRange(0, 1, 3, 1));
            Assert.Equal("=A2*2", editor.ViewModel.Rows[1].Cells[1].Text);
            Assert.Equal("=A3*2", editor.ViewModel.Rows[2].Cells[1].Text);
            Stopwatch timeout = Stopwatch.StartNew();
            while (editor.ViewModel.Rows[2].Cells[1].CachedNumericValue != 8 && timeout.Elapsed < TimeSpan.FromSeconds(5))
            {
                await Task.Delay(5);
            }
            Assert.Equal(6, editor.ViewModel.Rows[1].Cells[1].CachedNumericValue);
            Assert.Equal(8, editor.ViewModel.Rows[2].Cells[1].CachedNumericValue);
            using MemoryStream saved = new();
            await editor.SaveXlsxAsync(saved);
            saved.Position = 0;
            SheetData reopened = XlsxFile.Read(saved);
            Assert.Equal(new double?[] { 2, 6, 8 }, reopened.Cells.Select(row => row[1].FormulaNumericValue));
        }
        finally
        {
            editor.NewSheet(1, 1);
        }
    }

    private static EditorControl CreateEditor()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = Enumerable.Range(0, 8).Select(_ => Enumerable.Repeat(CellData.Plain(string.Empty), 7).ToArray()).ToArray(),
            ColumnWidths = Enumerable.Repeat(80d, 7).ToArray(),
            RowHeights = Enumerable.Repeat(24d, 8).ToArray(),
        });
        return editor;
    }

    private static Border Handle(SpreadsheetGrid grid) =>
        grid.GetVisualDescendants().OfType<Border>().Single(border => border.Name == "FillHandle");
}
