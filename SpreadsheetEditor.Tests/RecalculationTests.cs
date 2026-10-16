using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;

using ClosedXML.Excel;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Formulas;
using SpreadsheetEditor.Controls.Models;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class RecalculationTests
{
    [Fact]
    public async Task CpuWorkerUsesBelowNormalPriorityWithoutChangingTheCallingThread()
    {
        Type worker = typeof(EditorControl).Assembly.GetType("SpreadsheetEditor.Controls.BackgroundWork", throwOnError: true)!;
        MethodInfo run = worker.GetMethod("Run", BindingFlags.Static | BindingFlags.NonPublic)!.MakeGenericMethod(typeof(ThreadPriority));
        ThreadPriority priority = Thread.CurrentThread.Priority;
        var task = (Task<ThreadPriority>)run.Invoke(null, new object[]
        {
            new Func<ThreadPriority>(() => Thread.CurrentThread.Priority), CancellationToken.None,
        })!;
        Assert.Equal(ThreadPriority.BelowNormal, await task);
        Assert.Equal(priority, Thread.CurrentThread.Priority);
    }

    [AvaloniaFact]
    public async Task RecalculationKeepsVisibleCellsAndChartsOnScreenWhileDispatcherRuns()
    {
        const int rows = 20000, formulas = 30;
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Summary", new SheetData
            {
                Cells = Enumerable.Range(0, formulas).Select(index => new[]
                {
                    CellData.Plain($"Category {index}"),
                    CellData.Plain($"=COUNTIF(Data!A1:A{rows},\"ready\")") with
                    { FormulaResult = "-1", FormulaNumericValue = -1, NumberFormat = "#,##0" },
                }).ToArray(),
                Charts = Enumerable.Range(0, 4).Select(index => new SheetChart
                {
                    Kind = ChartKind.Line,
                    Series = new[] { new ChartSeriesSpec("Count", new CellRange(0, 1, formulas, 1)) },
                    Categories = new CellRange(0, 0, formulas, 1),
                    AnchorRow = index * 7, AnchorColumn = 3, Width = 300, Height = 130,
                }).ToArray(),
            }),
            ("Data", new SheetData { Cells = Enumerable.Range(0, rows).Select(_ => new[] { CellData.Plain("ready") }).ToArray() }),
        });
        Window window = new() { Content = editor, Width = 1000, Height = 800 };
        window.Show();
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Stopwatch timeout = Stopwatch.StartNew();
        DispatcherTimer timer = new() { Interval = TimeSpan.FromMilliseconds(5) };
        try
        {
            while (grid.GetVisualDescendants().OfType<Border>().Count(border => border.Name == "ChartCard") < 4 &&
                   timeout.Elapsed < TimeSpan.FromSeconds(5)) await Task.Delay(10);
            Border card = grid.GetVisualDescendants().OfType<Border>().First(border => border.Name == "ChartCard");
            Canvas overlay = Assert.IsType<Canvas>(card.Parent);
            int minimumChartCount = overlay.Children.Count;
            Assert.Equal(4, minimumChartCount);
            overlay.Children.CollectionChanged += (_, _) => minimumChartCount = Math.Min(minimumChartCount, overlay.Children.Count);
            TextBlock cell = grid.GetVisualDescendants().OfType<TextBlock>().First(block =>
                block.Inlines?.OfType<Run>().Any(run => run.Text == "-1") == true);
            int ticks = 0;
            TimeSpan previous = timeout.Elapsed, longestGap = TimeSpan.Zero;
            timer.Tick += (_, _) =>
            {
                TimeSpan now = timeout.Elapsed;
                longestGap = now - previous > longestGap ? now - previous : longestGap;
                previous = now;
                ticks++;
            };
            timer.Start();
            editor.Sheets[1].SetCell(new CellAddress(0, 0), "changed");
            while ((editor.Sheets[0].Rows[formulas - 1].Cells[1].CachedNumericValue != rows - 1 ||
                    cell.Inlines?.OfType<Run>().Any(run => run.Text == "19,999") != true) &&
                   timeout.Elapsed < TimeSpan.FromSeconds(15)) await Task.Delay(5);
            await Task.Delay(200); // include the chart refresh, not just formula evaluation
            Assert.Equal(rows - 1, editor.Sheets[0].Rows[formulas - 1].Cells[1].CachedNumericValue);
            Assert.Contains(cell, grid.GetVisualDescendants());
            Assert.Contains(cell.Inlines!.OfType<Run>(), run => run.Text == "19,999");
            Assert.Equal(4, minimumChartCount);
            Assert.True(ticks > 10);
            Assert.True(longestGap < TimeSpan.FromMilliseconds(500), $"Dispatcher was blocked for {longestGap}.");
        }
        finally
        {
            timer.Stop();
            editor.NewSheet(1, 1);
            window.Close();
            await Task.Delay(30);
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
        }
    }

    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BulkEditsAndUndoUseBackgroundCalculation(bool undo)
    {
        const int rows = 20000;
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Summary", new SheetData { Cells = Enumerable.Range(0, 100).Select(_ => new[]
            {
                CellData.Plain($"=COUNTIF(Data!A1:A{rows},\"ready\")") with { FormulaResult = "123", FormulaNumericValue = 123 },
            }).ToArray() }),
            ("Data", new SheetData { Cells = Enumerable.Range(0, rows).Select(_ => new[] { CellData.Plain("ready") }).ToArray() }),
        });
        var data = editor.Sheets[1];
        if (undo)
        {
            data.Checkpoint();
            data.SetCell(new CellAddress(0, 0), "changed");
        }
        Stopwatch clock = Stopwatch.StartNew();
        if (undo) data.Undo();
        else data.RunBatch(() => data.SetCell(new CellAddress(0, 0), "changed"));
        Assert.Equal(123, editor.Sheets[0].Rows[^1].Cells[0].CachedNumericValue);
        int heartbeats = 0;
        DispatcherTimer timer = new() { Interval = TimeSpan.FromMilliseconds(5) };
        timer.Tick += (_, _) => heartbeats++;
        timer.Start();
        try
        {
            double expected = undo ? rows : rows - 1;
            while (editor.Sheets[0].Rows[^1].Cells[0].CachedNumericValue != expected && clock.Elapsed < TimeSpan.FromSeconds(15))
                await Task.Delay(5);
            Assert.Equal(expected, editor.Sheets[0].Rows[^1].Cells[0].CachedNumericValue);
            Assert.True(heartbeats > 10);
        }
        finally { timer.Stop(); editor.NewSheet(1, 1); }
    }

    [AvaloniaFact]
    public async Task EditingReferencedValueKeepsDispatcherResponsiveDuringRangeCalculations()
    {
        const int rows = 20000;
        const int formulas = 100;
        CellData[][] summary = Enumerable.Range(0, formulas)
            .Select(_ => new[] { CellData.Plain(string.Empty), CellData.Plain($"=COUNTIF(Data!A1:A{rows},\"ready\")") })
            .ToArray();
        summary[16][0] = CellData.Plain("Note below the table");
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Monthly Stats", new SheetData { Cells = summary }),
            ("Data", new SheetData
            {
                Cells = Enumerable.Range(0, rows).Select(_ => new[] { CellData.Plain("ready") }).ToArray(),
            }),
        });

        Stopwatch clock = Stopwatch.StartNew();
        TimeSpan previousTick = clock.Elapsed;
        TimeSpan longestGap = TimeSpan.Zero;
        int ticks = 0;
        DispatcherTimer timer = new() { Interval = TimeSpan.FromMilliseconds(5) };
        timer.Tick += (_, _) =>
        {
            TimeSpan now = clock.Elapsed;
            longestGap = now - previousTick > longestGap ? now - previousTick : longestGap;
            previousTick = now;
            ticks++;
        };
        timer.Start();
        try
        {
            editor.Sheets[1].SetCell(new CellAddress(0, 0), "changed");
            while (editor.ViewModel.Rows[^1].Cells[1].CachedNumericValue != rows - 1 &&
                   clock.Elapsed < TimeSpan.FromSeconds(15))
            {
                await Task.Delay(5);
            }

            Assert.Equal(rows - 1, editor.ViewModel.Rows[^1].Cells[1].CachedNumericValue);
            Assert.True(ticks > 10, "Input/render work must continue while formulas are being calculated.");
            Assert.True(longestGap < TimeSpan.FromMilliseconds(500), $"Dispatcher was blocked for {longestGap}.");
            Assert.Equal("changed", editor.Sheets[1].Rows[0].Cells[0].Text);
        }
        finally
        {
            timer.Stop();
        }
    }

    [AvaloniaFact]
    public async Task EditingUnreferencedNotePreservesCachedTotals()
    {
        CellData[][] cells = Enumerable.Range(0, 17)
            .Select(_ => new[] { CellData.Plain(string.Empty), CellData.Plain(string.Empty) }).ToArray();
        cells[0][1] = CellData.Plain("=SUM(Data!A1:A2)") with { FormulaResult = "123", FormulaNumericValue = 123 };
        cells[16][0] = CellData.Plain("Original note");
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Monthly Stats", new SheetData { Cells = cells }),
            ("Data", new SheetData { Cells = new[] { new[] { CellData.Plain("1") }, new[] { CellData.Plain("2") } } }),
        });
        editor.ViewModel.SetCell(new CellAddress(16, 0), "Updated note");
        await Task.Delay(650);
        Assert.Equal("Updated note", editor.ViewModel.Rows[16].Cells[0].Text);
        Assert.Equal(123, editor.ViewModel.Rows[0].Cells[1].CachedNumericValue);
    }

    [AvaloniaFact]
    public async Task UnopenedDependentSheetRecalculatesWhenItIsLoadedAfterAnEdit()
    {
        using XLWorkbook workbook = new();
        workbook.Worksheets.Add("Data").Cell(1, 1).Value = 1;
        workbook.Worksheets.Add("Summary").Cell(1, 1).FormulaA1 = "Data!A1*2";
        workbook.RecalculateAllFormulas();
        using MemoryStream source = new();
        workbook.SaveAs(source);
        source.Position = 0;
        EditorControl editor = new();
        await editor.OpenXlsxAsync(source);
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Stopwatch timeout = Stopwatch.StartNew();
        while (grid.ReadOnly && timeout.Elapsed < TimeSpan.FromSeconds(10))
        {
            await Task.Delay(10);
        }
        Assert.False(grid.ReadOnly);
        editor.ViewModel.SetCell(new CellAddress(0, 0), "3");
        StackPanel tabs = editor.FindControl<StackPanel>("SheetTabs")!;
        ((Button)tabs.Children[1]).RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        while ((editor.ViewModel.Name != "Summary" || grid.ReadOnly ||
                editor.ViewModel.Rows[0].Cells[0].CachedNumericValue != 6) && timeout.Elapsed < TimeSpan.FromSeconds(10))
        {
            await Task.Delay(10);
        }
        Assert.Equal("Summary", editor.ViewModel.Name);
        Assert.Equal(6, editor.ViewModel.Rows[0].Cells[0].CachedNumericValue);
    }

    [AvaloniaFact]
    public async Task NewEditsSupersedePendingCrossSheetResults()
    {
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Summary", new SheetData { Cells = new[] { new[] { CellData.Plain("=Data!A1*2") } } }),
            ("Data", new SheetData { Cells = new[] { new[] { CellData.Plain("1") } } }),
        });
        editor.Sheets[1].SetCell(new CellAddress(0, 0), "2");
        await Task.Delay(305);
        editor.Sheets[1].SetCell(new CellAddress(0, 0), "3");
        Stopwatch timeout = Stopwatch.StartNew();
        while (editor.Sheets[0].Rows[0].Cells[0].CachedNumericValue != 6 && timeout.Elapsed < TimeSpan.FromSeconds(5))
        {
            await Task.Delay(5);
        }
        Assert.Equal(6, editor.Sheets[0].Rows[0].Cells[0].CachedNumericValue);
        Assert.Equal("3", editor.Sheets[1].Rows[0].Cells[0].Text);
    }

    [AvaloniaFact]
    public async Task NamedRangeInputsStillTriggerCrossSheetCalculation()
    {
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Summary", new SheetData
            {
                Cells = new[] { new[] { CellData.Plain("=SUM(TotalInputs)") } },
                Names = new Dictionary<string, string> { ["TotalInputs"] = "'Data'!A1:A2" },
            }),
            ("Data", new SheetData { Cells = new[] { new[] { CellData.Plain("1") }, new[] { CellData.Plain("2") } } }),
        });
        editor.Sheets[1].SetCell(new CellAddress(0, 0), "5");
        Stopwatch timeout = Stopwatch.StartNew();
        while (editor.Sheets[0].Rows[0].Cells[0].CachedNumericValue != 7 && timeout.Elapsed < TimeSpan.FromSeconds(5))
        {
            await Task.Delay(5);
        }
        Assert.Equal(7, editor.Sheets[0].Rows[0].Cells[0].CachedNumericValue);
    }

    [AvaloniaFact]
    public async Task StartingNewWorkbookDiscardsPendingResults()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[] { CellData.Plain("1"), CellData.Plain("=A1*2") } },
        });
        editor.ViewModel.SetCell(new CellAddress(0, 0), "5");
        await Task.Delay(305);
        editor.NewSheet(1, 1);
        editor.ViewModel.SetCell(new CellAddress(0, 0), "new workbook");
        await Task.Delay(350);
        Assert.Equal("new workbook", editor.ViewModel.Rows[0].Cells[0].Text);
        Assert.False(editor.ViewModel.Rows[0].Cells[0].IsFormula);
        Assert.Null(editor.ViewModel.Rows[0].Cells[0].ComputedValue);
    }

    [Fact]
    public void ImportedDateTextCanBeUsedByDateFormulas()
    {
        FormulaEngine engine = new(address => address.Column switch
        {
            0 => "2026-10-15 14:30:00",
            1 => "=YEAR(A1)",
            2 => "=MONTH(A1)",
            _ => "=DAY(A1)",
        });
        Assert.Equal(2026, engine.Evaluate(new CellAddress(0, 1)).Number);
        Assert.Equal(10, engine.Evaluate(new CellAddress(0, 2)).Number);
        Assert.Equal(15, engine.Evaluate(new CellAddress(0, 3)).Number);
    }

    [AvaloniaFact]
    public async Task VolatileFormulasStillUpdateWhenAnUnreferencedCellChanges()
    {
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Clock", new SheetData
            {
                Cells = new[] { new[] { CellData.Plain("=TODAY()") with { FormulaResult = "1", FormulaNumericValue = 1 } } },
            }),
            ("Notes", new SheetData { Cells = new[] { new[] { CellData.Plain("note") } } }),
        });
        editor.Sheets[1].SetCell(new CellAddress(0, 0), "updated note");
        Stopwatch timeout = Stopwatch.StartNew();
        while (editor.Sheets[0].Rows[0].Cells[0].CachedNumericValue == 1 && timeout.Elapsed < TimeSpan.FromSeconds(5))
        {
            await Task.Delay(5);
        }
        Assert.Equal(Math.Floor(DateTime.Today.ToOADate()), editor.Sheets[0].Rows[0].Cells[0].CachedNumericValue);
    }
}
