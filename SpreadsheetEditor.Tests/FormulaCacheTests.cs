using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.Threading;

using ClosedXML.Excel;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Formulas;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class FormulaCacheTests
{
    [AvaloniaFact]
    public async Task CachedNumbersTextBooleansErrorsAndEmptyStringsSurviveRepeatedSaves()
    {
        string[] formulas = { "12345678.12345678", "0", "-0.12345678", "TRUE", "FALSE", "\"TRUE\"", "\"42\"", "\"A < B & Ω\"", "\"\"", "1/0" };
        using XLWorkbook original = new();
        IXLWorksheet sheet = original.Worksheets.Add("Cached results");
        for (int row = 0; row < formulas.Length; row++)
        {
            sheet.Cell(row + 1, 1).FormulaA1 = formulas[row];
        }
        sheet.Cell(1, 1).Style.NumberFormat.Format = "#,##0.00";
        sheet.Cell(3, 1).Style.NumberFormat.Format = "0.0%";
        original.RecalculateAllFormulas();
        using MemoryStream source = new();
        original.SaveAs(source, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        source.Position = 0;
        // Compare with what the source file actually stores (ClosedXML rounds its own numeric caches).
        using XLWorkbook serialized = new(source);
        sheet = serialized.Worksheet(1);
        source.Position = 0;
        SheetData expectedData = XlsxFile.Read(source);
        source.Position = 0;

        EditorControl editor = new();
        await editor.OpenXlsxAsync(source);
        for (int row = 0; row < formulas.Length; row++)
        {
            Assert.Equal(sheet.Cell(row + 1, 1).CachedValue.Type == XLDataType.Number,
                editor.ViewModel.Rows[row].Cells[0].CachedNumericValue is not null);
            Assert.Equal(expectedData.Cells[row][0].FormulaResult, editor.ViewModel.Rows[row].Cells[0].ComputedValue);
        }
        using MemoryStream first = new();
        await editor.SaveXlsxAsync(first);
        await WaitForActiveSheetLoad(editor);
        first.Position = 0;
        EditorControl fullyLoaded = new();
        fullyLoaded.LoadWorkbook(XlsxFile.ReadWorkbook(first));
        using MemoryStream second = new();
        fullyLoaded.SaveXlsx(second);
        second.Position = 0;
        using XLWorkbook reopened = new(second);
        for (int row = 0; row < formulas.Length; row++)
        {
            IXLCell expected = sheet.Cell(row + 1, 1);
            IXLCell actual = reopened.Worksheet(1).Cell(row + 1, 1);
            Assert.Equal(expected.FormulaA1, actual.FormulaA1);
            Assert.Equal(expected.CachedValue.Type, actual.CachedValue.Type);
            Assert.Equal(expected.CachedValue, actual.CachedValue);
        }
    }

    [AvaloniaTheory]
    [InlineData("_x0041_")]
    [InlineData("A\u0001B")]
    [InlineData("First line\r\nSecond line & Ω")]
    public async Task CachedFormulaTextPreservesEscapesAndControlCharacters(string text)
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[]
            {
                new[] { CellData.Plain("=UNSUPPORTED()") with { FormulaResult = text, FormulaValue = CellValue.FromText(text) } },
            },
        });
        using MemoryStream saved = new();
        editor.SaveXlsx(saved);
        saved.Position = 0;
        using (XLWorkbook workbook = new(saved))
        {
            // ClosedXML leaves OOXML escape sequences in cached formula strings undecoded.
            Assert.Equal(XLDataType.Text, workbook.Worksheet(1).Cell(1, 1).CachedValue.Type);
        }
        saved.Position = 0;
        EditorControl reopened = new();
        await reopened.OpenXlsxAsync(saved);
        Assert.Equal(text, reopened.ViewModel.Rows[0].Cells[0].ComputedValue);
        await WaitForActiveSheetLoad(reopened);
        saved.Position = 0;
        SheetData data = XlsxFile.Read(saved);
        Assert.Equal(text, data.Cells[0][0].FormulaResult);
        reopened.LoadSheet(data);
        using MemoryStream savedAgain = new();
        reopened.SaveXlsx(savedAgain);
        savedAgain.Position = 0;
        Assert.Equal(text, XlsxFile.Read(savedAgain).Cells[0][0].FormulaResult);
    }

    [AvaloniaFact]
    public async Task SavingUnopenedWorksheetPreservesItsCachedFormulaResults()
    {
        using XLWorkbook original = new();
        original.Worksheets.Add("Data").Cell(1, 1).Value = 7;
        original.Worksheets.Add("Unopened").Cell(1, 1).FormulaA1 = "Data!A1*3";
        original.RecalculateAllFormulas();
        using MemoryStream source = new();
        original.SaveAs(source, new SaveOptions { EvaluateFormulasBeforeSaving = true });
        source.Position = 0;
        EditorControl editor = new();
        await editor.OpenXlsxAsync(source);
        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        await WaitForActiveSheetLoad(editor);
        saved.Position = 0;
        using XLWorkbook reopened = new(saved);
        Assert.Equal(21, reopened.Worksheet("Unopened").Cell(1, 1).CachedValue.GetNumber());
    }

    [AvaloniaFact]
    public void FormulaCachesSurviveChartSerialization()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[]
            {
                new[] { CellData.Plain("=2*3") with { FormulaResult = "6", FormulaNumericValue = 6 } },
                new[] { CellData.Plain("=2*4") with { FormulaResult = "8", FormulaNumericValue = 8 } },
            },
            Charts = new[]
            {
                new SheetChart
                {
                    Kind = ChartKind.Line,
                    Series = new[] { new ChartSeriesSpec("Values", new CellRange(0, 0, 2, 1)) },
                    AnchorRow = 0,
                    AnchorColumn = 2,
                },
            },
        });
        using MemoryStream saved = new();
        editor.SaveXlsx(saved);
        saved.Position = 0;
        SheetData read = XlsxFile.Read(saved);
        Assert.Single(read.Charts);
        Assert.Equal(new double?[] { 6, 8 }, read.Cells.Select(row => row[0].FormulaNumericValue));
    }

    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CrossSheetResultsSurviveSavingAndBothReopenStages(bool asynchronousSave)
    {
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Summary", new SheetData
            {
                Cells = new[] { new[] { CellData.Plain("=Data!A1/3") with { NumberFormat = "0.00%" } } },
            }),
            ("Data", new SheetData { Cells = new[] { new[] { CellData.Plain("1") } } }),
        });
        editor.Sheets[1].SetCell(new CellAddress(0, 0), "2");
        Stopwatch timeout = Stopwatch.StartNew();
        while (editor.Sheets[0].Rows[0].Cells[0].CachedNumericValue != 2d / 3 && timeout.Elapsed < TimeSpan.FromSeconds(5))
        {
            await Task.Delay(5);
        }
        Assert.Equal(2d / 3, editor.Sheets[0].Rows[0].Cells[0].CachedNumericValue);
        string display = editor.Sheets[0].Rows[0].Cells[0].DisplayText;

        using MemoryStream saved = new();
        if (asynchronousSave)
        {
            await editor.SaveXlsxAsync(saved);
        }
        else
        {
            editor.SaveXlsx(saved);
        }
        saved.Position = 0;
        using (XLWorkbook workbook = new(saved))
        {
            IXLCell cell = workbook.Worksheet("Summary").Cell(1, 1);
            Assert.Equal("Data!A1/3", cell.FormulaA1);
            Assert.Equal(XLDataType.Number, cell.CachedValue.Type);
            Assert.Equal(2d / 3, cell.CachedValue.GetNumber());
        }

        saved.Position = 0;
        EditorControl reopened = new();
        await reopened.OpenXlsxAsync(saved);
        Assert.Equal(display, reopened.ViewModel.Rows[0].Cells[0].DisplayText);
        Assert.Equal(2d / 3, reopened.ViewModel.Rows[0].Cells[0].CachedNumericValue);
        await WaitForActiveSheetLoad(reopened);
        Assert.Equal(display, reopened.ViewModel.Rows[0].Cells[0].DisplayText);
        Assert.Equal(2d / 3, reopened.ViewModel.Rows[0].Cells[0].CachedNumericValue);
    }

    private static async Task WaitForActiveSheetLoad(EditorControl editor)
    {
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Stopwatch timeout = Stopwatch.StartNew();
        while (grid.ReadOnly && timeout.Elapsed < TimeSpan.FromSeconds(5))
        {
            await Task.Delay(5);
        }
        Assert.False(grid.ReadOnly);
        // Let the load-completion dispatcher callbacks drain before the headless session ends.
        await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
    }
}
