using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Charts;
using Avalonia.Controls.Documents;
using Avalonia.Interactivity;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Formulas;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;
using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class PresentationTests
{
    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WorkbooksWithoutExplicitStyleTablesOpenInPreviewAndFullImport(bool emptyStylesPart)
    {
        using ClosedXML.Excel.XLWorkbook workbook = new();
        var sheet = workbook.Worksheets.Add("Plain");
        sheet.Cell(1, 1).Value = "Text";
        sheet.Cell(1, 2).Value = 42;
        using MemoryStream source = new(); workbook.SaveAs(source);
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(source, true))
        {
            var part = document.WorkbookPart!;
            if (emptyStylesPart) part.WorkbookStylesPart!.Stylesheet = new DocumentFormat.OpenXml.Spreadsheet.Stylesheet();
            else part.DeletePart(part.WorkbookStylesPart!);
            foreach (var worksheet in part.WorksheetParts)
            {
                foreach (var cell in worksheet.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>()) cell.StyleIndex = null;
                foreach (var column in worksheet.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Column>()) column.Style = null;
                foreach (var row in worksheet.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>()) row.StyleIndex = null;
            }
        }
        source.Position = 0;
        SheetData imported = XlsxFile.Read(source);
        Assert.Equal("Text", imported.Cells[0][0].Text);
        Assert.Equal("42", imported.Cells[0][1].Text);
        Assert.Equal(TextAlignment.Right, imported.Cells[0][1].Alignment);
        source.Position = 0;
        EditorControl editor = new();
        try
        {
            await editor.OpenXlsxAsync(source);
            Assert.Equal("42", editor.ViewModel.Rows[0].Cells[1].DisplayText);
            Assert.Equal(TextAlignment.Right, editor.ViewModel.Rows[0].Cells[1].TextAlignment);
        }
        finally { editor.NewSheet(1, 1); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle); }
    }

    [Theory]
    [InlineData(632.12, "#,##0.00 $", "632.12 $")]
    [InlineData(1234.5, "#,##0.00 [$€-40C]", "1,234.50 €")]
    [InlineData(-42.5, "#,##0.00;[Red]-#,##0.00", "-42.50")]
    [InlineData(123456789, "0.00E+00", "1.23E+08")]
    [InlineData(0.375, "# ?/?", " 3/8")]
    [InlineData(42, "00000", "00042")]
    [InlineData(5145551234, "(###) ###-####", "(514) 555-1234")]
    [InlineData(0, "#,##0;-#,##0;\"-\"", "-")]
    public void ExcelNumberFormatsKeepTheirLiteralsAndSections(double value, string code, string expected) =>
        Assert.Equal(expected, NumberFormatter.Format(value, code));

    [Fact]
    public void DateAndNegativeColourFormatsAreApplied()
    {
        Assert.Equal("2:35 PM", NumberFormatter.Format(new TimeSpan(14, 35, 0).TotalDays, "h:mm AM/PM"));
        Assert.Equal("2026-11-04", NumberFormatter.Format(new DateTime(2026, 11, 4).ToOADate(), "yyyy-mm-dd"));
        Assert.Equal(Colors.Red, NumberFormatter.Foreground(-915, "#,##0;[Red](#,##0)"));
        Assert.Null(NumberFormatter.Foreground(915, "#,##0;[Red](#,##0)"));
    }

    [AvaloniaFact]
    public void AdvancedFormattingVisibilityGroupsAndProtectionSurviveSave()
    {
        string[] styles = { "thin", "medium", "thick", "dashed", "dotted", "double", "hair", "mediumDashDot" };
        List<IReadOnlyList<CellData>> rows = new();
        foreach (string style in styles)
        {
            CellEdge edge = new(style, Colors.Black);
            rows.Add(new[] { CellData.Plain(style) with { Appearance = new CellAppearance
            {
                Left = edge, Right = edge, Top = edge, Bottom = edge,
                Rotation = 45, Indent = 3, ShrinkToFit = true, DoubleUnderline = true,
                Script = "superscript", Locked = false, VerticalAlignment = VerticalAlignment.Bottom,
                Fill = new CellFill { Pattern = "darkGrid", Foreground = Colors.Blue, Background = Colors.White },
            }, Underline = true } });
        }
        rows[1] = new[] { rows[1][0] with { Appearance = rows[1][0].Appearance! with { Fill = new CellFill { Stops = new[] { new CellGradientStop(0, Colors.White), new CellGradientStop(1, Colors.Blue) } } } } };
        SheetData data = new()
        {
            Cells = rows, HiddenColumns = new[] { 0 }, HiddenRows = new[] { 3, 4 },
            RowGroups = new[] { new SheetOutlineGroup(3, 2, 1, true, false) },
            ColumnGroups = new[] { new SheetOutlineGroup(0, 1, 1, true, true) }, ColumnWidths = new[] { 100d },
            ProtectionXml = "<sheetProtection xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" sheet=\"1\" password=\"ABCD\" />",
        };
        EditorControl editor = new();
        editor.LoadWorkbook(new[] { ("Visible", data), ("Lookup", new SheetData { Cells = new[] { new[] { CellData.Plain("hidden data") } }, Visibility = "hidden" }) });
        using MemoryStream saved = new(); editor.SaveXlsx(saved);
        saved.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(saved, false))
        {
            var errors = new OpenXmlValidator().Validate(document).ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(e => e.Path?.XPath + ": " + e.Description)));
        }
        saved.Position = 0;
        var reopened = XlsxFile.ReadWorkbook(saved);
        Assert.Equal("hidden", reopened[1].Data.Visibility);
        SheetData result = reopened[0].Data;
        Assert.Equal(styles, result.Cells.Select(row => row[0].Appearance!.Left!.Style));
        Assert.Equal(2, result.Cells[1][0].Appearance!.Fill!.Stops.Count);
        Assert.True(result.Cells[0][0].Appearance!.DoubleUnderline);
        Assert.Equal("superscript", result.Cells[0][0].Appearance!.Script);
        Assert.Equal(45, result.Cells[0][0].Appearance!.Rotation);
        Assert.False(result.Cells[0][0].Appearance!.Locked);
        Assert.Contains(0, result.HiddenColumns);
        Assert.Equal(data.RowGroups, result.RowGroups);
        Assert.Contains("ABCD", result.ProtectionXml);
    }

    [AvaloniaFact]
    public async Task ProtectedCellsBlockTypingButUnlockedInputsRemainEditable()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[] { CellData.Plain("locked"), CellData.Plain("input") with { Appearance = new CellAppearance { Locked = false } } } },
            ProtectionXml = "<sheetProtection xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" sheet=\"1\" />",
        });
        Window window = new() { Content = editor, Width = 800, Height = 500 }; window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            grid.SelectCell(new CellAddress(0, 0)); grid.FocusForSelection();
            window.KeyTextInput("oops");
            Assert.Equal("locked", editor.ViewModel.GetCell(new CellAddress(0, 0))!.Text);
            grid.SelectCell(new CellAddress(0, 1)); grid.FocusForSelection();
            window.KeyTextInput("ok");
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            Assert.Equal("ok", editor.ViewModel.GetCell(new CellAddress(0, 1))!.Text);
            Assert.False(editor.ViewModel.FillCells(new CellRange(0, 1, 1, 1), new CellRange(0, 0, 1, 2)));
        }
        finally { editor.NewSheet(1, 1); window.Close(); }
    }

    [AvaloniaFact]
    public void ConditionalTextExpressionsAndIconsRemainLiveAfterSaving()
    {
        CellRange range = new(1, 0, 3, 1);
        SheetCellValueRule[] rules =
        {
            new() { Range = range with { Column = 1 }, Threshold = 0, Comparison = CellValueComparison.Equal, TextThreshold = "Blocked", Foreground = Colors.Red, Background = Colors.Pink, Priority = 1 },
            new() { Range = range with { Column = 1 }, Threshold = 0, Comparison = CellValueComparison.Equal, TextThreshold = "Done", Foreground = Colors.Green, Priority = 2 },
            new() { Range = range, Threshold = 0, Comparison = CellValueComparison.Equal, Expression = "$C2=\"P1\"", Bold = true, Priority = 3 },
            new() { Range = range with { Column = 3 }, Threshold = 0, Comparison = CellValueComparison.Equal, Expression = "COUNTIF($D$2:$D$4,D2)>1", Bold = true, Foreground = Colors.Red, Priority = 4 },
            new() { Range = range with { Column = 4 }, Threshold = 0, Comparison = CellValueComparison.Equal, IconSet = "3Arrows", IconThresholds = new[] { -100d, 0, 1 }, Priority = 5 },
        };
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[]
            {
                new[] { "Ticket", "Status", "Priority", "Duplicate", "Trend" },
                new[] { "1", "Blocked", "P1", "alpha", "-1" },
                new[] { "2", "Done", "P2", "alpha", "0" },
                new[] { "3", "Open", "P1", "beta", "10" },
            }.Select(row => row.Select(CellData.Plain).ToArray()).ToArray(), CellValueRules = rules,
        });
        using MemoryStream saved = new(); editor.SaveXlsx(saved); saved.Position = 0;
        SheetData reopened = XlsxFile.Read(saved); editor.LoadSheet(reopened);
        Assert.Equal(5, editor.ViewModel.CellValueRules.Count);
        Assert.True(editor.ViewModel.MatchesRule(reopened.CellValueRules[0], 1, 1));
        Assert.True(editor.ViewModel.MatchesRule(reopened.CellValueRules[1], 2, 1));
        Assert.True(editor.ViewModel.MatchesRule(reopened.CellValueRules[2], 1, 0));
        Assert.True(editor.ViewModel.MatchesRule(reopened.CellValueRules[3], 1, 3));
        editor.ViewModel.SetCell(new CellAddress(1, 2), "P2");
        editor.ViewModel.SetCell(new CellAddress(1, 3), "gamma");
        Assert.False(editor.ViewModel.MatchesRule(reopened.CellValueRules[2], 1, 0));
        Assert.False(editor.ViewModel.MatchesRule(reopened.CellValueRules[3], 1, 3));
        Assert.Equal(new[] { -100d, 0, 1 }, reopened.CellValueRules[4].IconThresholds);
    }

    [AvaloniaFact]
    public async Task HiddenTabsStayHiddenDuringLazyLoadAndOutlineButtonsExpandRows()
    {
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Visible", new SheetData { Cells = Enumerable.Range(0, 8).Select(i => new[] { CellData.Plain(i.ToString()) }).ToArray(),
                HiddenRows = new[] { 3, 4 }, RowGroups = new[] { new SheetOutlineGroup(3, 2, 1, true, false) } }),
            ("Lookup", new SheetData { Cells = new[] { new[] { CellData.Plain("lookup") } }, Visibility = "hidden" }),
        });
        using MemoryStream saved = new(); editor.SaveXlsx(saved); saved.Position = 0;
        await editor.OpenXlsxAsync(saved);
        Assert.Single(editor.FindControl<StackPanel>("SheetTabs")!.Children);
        Window window = new() { Content = editor, Width = 800, Height = 600 }; window.Show();
        try
        {
            await Task.Delay(100);
            window.UpdateLayout(); window.CaptureRenderedFrame();
            Button toggle = editor.GetVisualDescendants().OfType<Button>().Single(b => b.Name == "RowOutlineToggle");
            toggle.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            Assert.False(editor.ViewModel.Rows[3].IsHidden);
            Assert.False(editor.ViewModel.Rows[4].IsHidden);
            Assert.False(editor.ViewModel.RowGroups[0].Collapsed);
        }
        finally { editor.NewSheet(1, 1); window.Close(); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle); }
    }

    [Fact]
    public void ScatterAndComboChartsReadEverySeriesAndUseAxisIds()
    {
        XNamespace c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        XElement Series(string valueElement, string categoryElement) => new(c + "ser", new XElement(c + categoryElement,
            new XElement(c + "numRef", new XElement(c + "f", "Data!A2:A13"))), new XElement(c + valueElement,
            new XElement(c + "numRef", new XElement(c + "f", "Data!B2:B13"))));
        XDocument scatter = new(new XElement(c + "chartSpace", new XElement(c + "chart", new XElement(c + "plotArea", new XElement(c + "scatterChart", Series("yVal", "xVal"))))));
        var parsed = XlsxCharts.ParseChartPart(scatter);
        Assert.Equal(ChartKind.Scatter, parsed.Kind);
        Assert.Equal("Data!B2:B13", Assert.Single(parsed.Series).ValueFormula);
        Assert.Equal("Data!A2:A13", parsed.Series[0].CategoryFormula);
        XDocument combo = new(new XElement(c + "chartSpace", new XElement(c + "chart", new XElement(c + "plotArea",
            new XElement(c + "barChart", Series("val", "cat"), new XElement(c + "axId", new XAttribute("val", 10)), new XElement(c + "axId", new XAttribute("val", 100))),
            new XElement(c + "lineChart", Series("val", "cat"), new XElement(c + "axId", new XAttribute("val", 10)), new XElement(c + "axId", new XAttribute("val", 200))),
            new XElement(c + "valAx", new XElement(c + "axId", new XAttribute("val", 100)), new XElement(c + "axPos", new XAttribute("val", "l"))),
            new XElement(c + "valAx", new XElement(c + "axId", new XAttribute("val", 200)), new XElement(c + "axPos", new XAttribute("val", "l")))))));
        var combined = XlsxCharts.ParseChartPart(combo);
        Assert.Equal(2, combined.Series.Count);
        Assert.False(combined.Series[0].SecondaryAxis);
        Assert.True(combined.Series[1].SecondaryAxis);
        Assert.Equal(ChartKind.Line, combined.Series[1].Kind);
    }

    [AvaloniaFact]
    public void ScatterAndMultipleChartsSaveValidCachesAndRadarUsesEveryCategory()
    {
        string[] categories = Enumerable.Range(1, 12).Select(i => i.ToString()).ToArray();
        ChartSeriesData[] series = { new("Values", Enumerable.Range(1, 12).Select(i => (double)i).ToArray()) };
        RadarChart radar = Assert.IsType<RadarChart>(ChartFactory.Build(ChartKind.Radar, categories, series, 400, 300));
        Assert.Equal(12, radar.AxisCount);
        Assert.Equal(categories, radar.AxisLabels);
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = categories.Select(s => new[] { CellData.Plain(s), CellData.Plain(s) }).ToArray(),
            Charts = Enumerable.Range(0, 3).Select(i => new SheetChart
            {
                Kind = i == 0 ? ChartKind.Scatter : ChartKind.Line,
                Series = new[] { new ChartSeriesSpec("Values", new CellRange(0, 1, 12, 1)) },
                Categories = new CellRange(0, 0, 12, 1), AnchorRow = i * 15, AnchorColumn = 3,
            }).ToArray(),
        });
        using MemoryStream saved = new(); editor.SaveXlsx(saved); saved.Position = 0;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(saved, false))
        {
            var errors = new OpenXmlValidator().Validate(document).ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(e => e.Description)));
        }
        saved.Position = 0;
        var charts = XlsxFile.Read(saved).Charts;
        Assert.Equal(3, charts.Count);
        Assert.Equal(ChartKind.Scatter, charts[0].Kind);
        Assert.Equal(series[0].Values, charts[0].Series[0].CachedValues);
    }

    [AvaloniaFact]
    public async Task LargeTableScrollingKeepsMenusLazyAndOnlyMaterializesTheViewport()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = Enumerable.Range(0, 5001).Select(row => Enumerable.Range(0, 10)
                .Select(column => CellData.Plain(row == 0 ? $"Column {column}" : $"Value {row}/{column}")).ToArray()).ToArray(),
            Tables = new[] { new SheetTable { Name = "Big", Range = new CellRange(0, 0, 5001, 10), AutoFilterRange = new CellRange(0, 0, 5001, 10) } },
        });
        Window window = new() { Content = editor, Width = 900, Height = 600 }; window.Show();
        try
        {
            window.UpdateLayout(); window.CaptureRenderedFrame();
            Border filter = editor.GetVisualDescendants().OfType<Border>().First(b => b.Name == "ColumnFilterButton");
            Assert.Empty(Assert.IsType<MenuFlyout>(filter.ContextFlyout).Items);
            ScrollViewer scroll = editor.GetVisualDescendants().OfType<ScrollViewer>().OrderByDescending(s => s.Extent.Width * s.Extent.Height).First();
            TextBlock retained = editor.GetVisualDescendants().OfType<TextBlock>().Single(t => t.Name == "CellText" && Equals(t.Tag, new CellAddress(12, 1)));
            scroll.Offset = new Vector(0, 200);
            window.UpdateLayout(); window.CaptureRenderedFrame();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
            Assert.Same(retained, editor.GetVisualDescendants().OfType<TextBlock>().Single(t => t.Name == "CellText" && Equals(t.Tag, new CellAddress(12, 1))));
            scroll.Offset = new Vector(0, 22000);
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
            window.UpdateLayout(); window.CaptureRenderedFrame();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
            TextBlock[] cells = editor.GetVisualDescendants().OfType<TextBlock>().Where(t => t.Name == "CellText").ToArray();
            Assert.InRange(cells.Length, 1, 1500);
            Assert.Contains(cells, t => t.Tag is CellAddress address && address.Row >= 900);
        }
        finally { editor.NewSheet(1, 1); window.Close(); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle); }
    }

    [AvaloniaFact]
    public void TablesApplyColumnBandsAndFirstColumnEmphasis()
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = Enumerable.Range(0, 5).Select(row => Enumerable.Range(0, 4).Select(col => CellData.Plain($"{row}/{col}")).ToArray()).ToArray(),
            Tables = new[] { new SheetTable { Name = "Banded", Range = new CellRange(0, 0, 5, 3),
                StyleName = "TableStyleDark10", ShowColumnStripes = true, EmphasizeFirstColumn = true } },
        });
        using MemoryStream saved = new(); editor.SaveXlsx(saved); saved.Position = 0;
        SheetData sheet = XlsxFile.Read(saved);
        Assert.True(sheet.Cells[1][0].Bold);
        Assert.NotNull(sheet.Cells[1][0].Background);
        Assert.NotEqual(sheet.Cells[1][0].Background, sheet.Cells[1][1].Background);
        Assert.Equal(sheet.Cells[1][0].Background, sheet.Cells[2][0].Background);
        Assert.Null(sheet.Cells[1][3].Background);
    }

    [Fact]
    public void ExcelRgbColoursIgnoreTheUnusedAlphaByte()
    {
        using ClosedXML.Excel.XLWorkbook workbook = new();
        var sheet = workbook.Worksheets.Add("Colours");
        sheet.Cell(1, 1).Value = "Header";
        sheet.Cell(1, 1).Style.Font.FontColor = ClosedXML.Excel.XLColor.FromArgb(0, 31, 78, 120);
        sheet.Cell(1, 1).Style.Fill.BackgroundColor = ClosedXML.Excel.XLColor.FromArgb(0, 255, 255, 0);
        using MemoryStream source = new(); workbook.SaveAs(source); source.Position = 0;
        CellData cell = XlsxFile.Read(source).Cells[0][0];
        Assert.Equal(Color.FromRgb(31, 78, 120), cell.Foreground);
        Assert.Equal(Colors.Yellow, cell.Background);
    }

    [AvaloniaFact]
    public void MissingFormulaCacheIsCalculatedRatherThanTreatedAsASavedBlank()
    {
        using ClosedXML.Excel.XLWorkbook workbook = new();
        var sheet = workbook.Worksheets.Add("Profit");
        sheet.Cell(1, 1).Value = 100;
        sheet.Cell(1, 2).Value = 1015;
        sheet.Cell(1, 3).FormulaA1 = "A1-B1";
        sheet.Cell(1, 3).Style.NumberFormat.Format = "#,##0;[Red](#,##0)";
        using MemoryStream source = new(); workbook.SaveAs(source); source.Position = 0;
        SheetData data = XlsxFile.Read(source);
        Assert.Null(data.Cells[0][2].FormulaResult);
        EditorControl editor = new(); editor.LoadSheet(data);
        Assert.Equal("(915)", editor.ViewModel.Rows[0].Cells[2].DisplayText);
        Assert.Equal(-915, editor.ViewModel.Rows[0].Cells[2].NumericValue);
    }
}
