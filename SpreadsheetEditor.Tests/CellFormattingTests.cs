namespace SpreadsheetEditor.Tests;

using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Input.Raw;
using Avalonia.Media;
using Avalonia.Threading;
using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.ViewModels;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;
using Xunit;

public sealed class CellFormattingTests
{
    [Theory]
    [InlineData(CellMergeMode.Merge)]
    [InlineData(CellMergeMode.Center)]
    [InlineData(CellMergeMode.Across)]
    public void MergeModesKeepTheirAnchorsAndUndoAsOneOperation(CellMergeMode mode)
    {
        SpreadsheetViewModel sheet = new(6, 6);
        sheet.SetCell(new CellAddress(1, 0), "First");
        sheet.SetCell(new CellAddress(2, 0), "Second");
        sheet.SetCell(new CellAddress(3, 0), "Third");
        sheet.Rows[1].Cells[0].Bold = true;
        Assert.True(sheet.MergeCells(new CellRange(1, 0, 3, 3), mode, discardContents: true));
        Assert.Equal(mode == CellMergeMode.Across ? 3 : 1, sheet.Merges.Count);
        Assert.Equal("First", sheet.Rows[1].Cells[0].Text);
        Assert.Equal(mode == CellMergeMode.Across ? "Second" : "", sheet.Rows[2].Cells[0].Text);
        Assert.Equal(mode == CellMergeMode.Center ? TextAlignment.Center : TextAlignment.Left, sheet.Rows[1].Cells[0].TextAlignment);
        Assert.True(sheet.Rows[1].Cells[0].Bold);
        sheet.Undo();
        Assert.Empty(sheet.Merges);
        Assert.Equal("Second", sheet.Rows[2].Cells[0].Text);
        Assert.False(sheet.CanUndo);
        sheet.Redo();
        Assert.Equal(mode == CellMergeMode.Across ? 3 : 1, sheet.Merges.Count);
    }

    [Fact]
    public void MergeDoesNotSilentlyDiscardValuesOrFormulas()
    {
        SpreadsheetViewModel sheet = new(3, 3);
        sheet.SetCell(new CellAddress(0, 0), "Keep");
        sheet.SetCell(new CellAddress(0, 1), "=1+1");
        Assert.True(sheet.MergeWouldDiscardData(new CellRange(0, 0, 2, 2), CellMergeMode.Merge));
        Assert.False(sheet.MergeCells(new CellRange(0, 0, 2, 2)));
        Assert.Equal("=1+1", sheet.Rows[0].Cells[1].Text);
        Assert.Empty(sheet.Merges);
        Assert.False(sheet.CanUndo);
    }

    [Fact]
    public void UnmergeFromAnyCoveredCellSplitsTheEntireMergeAndIsUndoable()
    {
        SpreadsheetViewModel sheet = new(5, 5);
        sheet.SetMerges(new[] { new MergeRange(0, 0, 2, 2), new MergeRange(3, 3, 2, 2) });
        sheet.SetCell(new CellAddress(0, 0), "Keep");
        Assert.True(sheet.UnmergeCells(new CellRange(1, 1, 1, 1)));
        Assert.Equal(new MergeRange(3, 3, 2, 2), Assert.Single(sheet.Merges));
        Assert.Equal("Keep", sheet.Rows[0].Cells[0].Text);
        sheet.Undo();
        Assert.Equal(2, sheet.Merges.Count);
    }

    [Fact]
    public void OverlappingMergesExpandToAValidNonOverlappingRegion()
    {
        SpreadsheetViewModel sheet = new(5, 5);
        sheet.SetMerges(new[] { new MergeRange(0, 0, 2, 2), new MergeRange(2, 1, 2, 2) });
        Assert.True(sheet.MergeCells(new CellRange(1, 1, 2, 2)));
        Assert.Equal(new MergeRange(0, 0, 4, 3), Assert.Single(sheet.Merges));
    }

    [Theory]
    [InlineData(CellBorderPreset.Outside)]
    [InlineData(CellBorderPreset.All)]
    [InlineData(CellBorderPreset.Inside)]
    [InlineData(CellBorderPreset.Top)]
    [InlineData(CellBorderPreset.Bottom)]
    [InlineData(CellBorderPreset.Left)]
    [InlineData(CellBorderPreset.Right)]
    public void BorderPresetsApplyToTheSelectedEdgesAndUndo(CellBorderPreset preset)
    {
        SpreadsheetViewModel sheet = new(4, 4);
        CellEdge edge = new("double", Colors.Red);
        sheet.Rows[1].Cells[1].Appearance = new CellAppearance { Rotation = 45, DoubleUnderline = true };
        Assert.True(sheet.ApplyBorders(new CellRange(1, 1, 2, 2), preset, edge));
        CellAppearance first = sheet.Rows[1].Cells[1].Appearance!;
        CellAppearance last = sheet.Rows[2].Cells[2].Appearance!;
        Assert.Equal(preset is CellBorderPreset.Outside or CellBorderPreset.All or CellBorderPreset.Left ? edge : null, first.Left);
        Assert.Equal(preset is CellBorderPreset.Outside or CellBorderPreset.All or CellBorderPreset.Top ? edge : null, first.Top);
        Assert.Equal(preset is CellBorderPreset.All or CellBorderPreset.Inside ? edge : null, first.Right);
        Assert.Equal(preset is CellBorderPreset.Outside or CellBorderPreset.All or CellBorderPreset.Bottom ? edge : null, last.Bottom);
        Assert.Equal(45, first.Rotation);
        Assert.True(first.DoubleUnderline);
        sheet.Undo();
        Assert.Null(sheet.Rows[1].Cells[1].Appearance!.Left);
        Assert.Null(sheet.Rows[2].Cells[2].Appearance);
        Assert.False(sheet.CanUndo);
        sheet.Redo();
        Assert.Equal(first, sheet.Rows[1].Cells[1].Appearance);
    }

    [Fact]
    public void RemovingBordersClearsBothSidesOfSharedEdges()
    {
        SpreadsheetViewModel sheet = new(3, 3);
        sheet.Rows[0].Cells[1].Appearance = new CellAppearance { Bottom = new CellEdge("thick", Colors.Blue) };
        sheet.ApplyBorders(new CellRange(1, 1, 1, 1), CellBorderPreset.All, new CellEdge("thin", Colors.Red));
        Assert.Equal(Colors.Red, sheet.Rows[0].Cells[1].Appearance!.Bottom!.Color);
        sheet.ApplyBorders(new CellRange(1, 1, 1, 1), CellBorderPreset.None, new CellEdge("thin", Colors.Black));
        Assert.Null(sheet.Rows[0].Cells[1].Appearance!.Bottom);
        Assert.Null(sheet.Rows[1].Cells[1].Appearance!.Top);
    }

    [Fact]
    public void ExistingBorderColorAndStyleCanBeChangedWithoutClearingOrAddingEdges()
    {
        SpreadsheetViewModel sheet = new(5, 5);
        CellRange range = new(1, 1, 2, 2);
        sheet.Rows[0].Cells[1].Appearance = new CellAppearance { Bottom = new CellEdge("thick", Colors.Red) };
        sheet.ApplyBorders(range, CellBorderPreset.Outside, new CellEdge("thick", Colors.Red));
        sheet.Rows[1].Cells[1].Appearance = sheet.Rows[1].Cells[1].Appearance! with { Rotation = 45 };
        int recalculations = 0;
        sheet.DataChanged += () => recalculations++;

        Assert.True(sheet.UpdateBorderFormatting(range, color: Colors.Blue));
        Assert.Equal(new CellEdge("thick", Colors.Blue), sheet.Rows[1].Cells[1].Appearance!.Top);
        Assert.Equal(new CellEdge("thick", Colors.Blue), sheet.Rows[0].Cells[1].Appearance!.Bottom);
        Assert.Null(sheet.Rows[1].Cells[1].Appearance!.Right);
        Assert.Null(sheet.Rows[1].Cells[1].Appearance!.Bottom);
        Assert.Equal(45, sheet.Rows[1].Cells[1].Appearance!.Rotation);

        Assert.True(sheet.UpdateBorderFormatting(range, style: "dotted"));
        Assert.Equal(new CellEdge("dotted", Colors.Blue), sheet.Rows[1].Cells[1].Appearance!.Left);
        Assert.Equal(new CellEdge("dotted", Colors.Blue), sheet.Rows[2].Cells[2].Appearance!.Bottom);
        Assert.Equal(0, recalculations);
        sheet.Undo();
        Assert.Equal(new CellEdge("thick", Colors.Blue), sheet.Rows[1].Cells[1].Appearance!.Top);
        sheet.Undo();
        Assert.Equal(new CellEdge("thick", Colors.Red), sheet.Rows[1].Cells[1].Appearance!.Top);
    }

    [Fact]
    public void RestylingBlankOrProtectedCellsDoesNotCreateBordersOrUndoEntries()
    {
        SpreadsheetViewModel sheet = new(3, 3);
        Assert.False(sheet.UpdateBorderFormatting(new CellRange(0, 0, 2, 2), color: Colors.Red));
        Assert.False(sheet.CanUndo);
        Assert.Null(sheet.Rows[0].Cells[0].Appearance);
        sheet.ProtectionXml = "<sheetProtection sheet=\"1\"/>";
        Assert.False(sheet.UpdateBorderFormatting(new CellRange(0, 0, 2, 2), style: "thick"));
        Assert.False(sheet.CanUndo);
    }

    [Fact]
    public void ProtectedSheetsAndExcelTablesCannotBeMerged()
    {
        SpreadsheetViewModel sheet = new(4, 4);
        sheet.ProtectionXml = "<sheetProtection sheet=\"1\"/>";
        Assert.False(sheet.MergeCells(new CellRange(0, 0, 2, 2)));
        Assert.False(sheet.ApplyBorders(new CellRange(0, 0, 2, 2), CellBorderPreset.All, new CellEdge("thin", Colors.Red)));
        sheet.ProtectionXml = null;
        sheet.SetTables(new[] { new SheetTable { Name = "Table1", Range = new CellRange(0, 0, 3, 3) } });
        Assert.False(sheet.MergeCells(new CellRange(1, 1, 2, 2)));
        Assert.False(sheet.MergeCells(new CellRange(1, 1, 0, 2)));
        Assert.Empty(sheet.Merges);
        Assert.False(sheet.CanUndo);
    }

    [AvaloniaTheory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task NewlyCreatedMergesAndColoredBordersSurviveSavingAndReopening(bool merged, bool empty)
    {
        EditorControl editor = new();
        editor.NewSheet(6, 6);
        SpreadsheetViewModel sheet = editor.ViewModel;
        if (!empty) sheet.SetCell(new CellAddress(1, 0), "Ticket example");
        CellRange range = new(1, 0, 3, 3);
        if (merged) Assert.True(sheet.MergeCells(range, CellMergeMode.Center));
        Assert.True(sheet.ApplyBorders(range, CellBorderPreset.Outside, new CellEdge("double", Colors.Red)));
        using MemoryStream saved = new();
        await editor.SaveXlsxAsync(saved);
        saved.Position = 0;
        SheetData reopened = XlsxFile.Read(saved);
        Assert.Equal(merged ? 1 : 0, reopened.Merges.Count);
        Assert.Equal(Colors.Red, reopened.Cells[1][0].Appearance!.Left!.Color);
        Assert.Equal("double", reopened.Cells[3][2].Appearance!.Bottom!.Style);
        Assert.Equal(Colors.Red, reopened.Cells[1][2].Appearance!.Right!.Color);
    }

    [Fact]
    public void AllBordersTreatMergedCellsAsOneCellAndFormattingDoesNotRecalculate()
    {
        SpreadsheetViewModel sheet = new(4, 4);
        int recalculations = 0;
        sheet.DataChanged += () => recalculations++;
        Assert.True(sheet.MergeCells(new CellRange(0, 0, 2, 2)));
        sheet.ApplyBorders(new CellRange(0, 0, 1, 1), CellBorderPreset.All, new CellEdge("thin", Colors.Red));
        Assert.Null(sheet.Rows[0].Cells[0].Appearance!.Right);
        Assert.Null(sheet.Rows[0].Cells[0].Appearance!.Bottom);
        Assert.Equal(Colors.Red, sheet.Rows[0].Cells[1].Appearance!.Right!.Color);
        Assert.Equal(Colors.Red, sheet.Rows[1].Cells[0].Appearance!.Bottom!.Color);
        Assert.Equal(0, recalculations);
    }

    [AvaloniaFact]
    public async Task GridCommandsUseSelectionAndRespectReadOnly()
    {
        EditorControl editor = new();
        editor.NewSheet(5, 5);
        SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
        Window window = new() { Content = editor, Width = 1000, Height = 600 };
        window.Show();
        try
        {
            grid.SelectCell(new CellAddress(1, 0));
            grid.FocusForSelection();
            window.KeyPress(Key.Right, RawInputModifiers.Shift, PhysicalKey.ArrowRight, null);
            window.KeyPress(Key.Right, RawInputModifiers.Shift, PhysicalKey.ArrowRight, null);
            window.KeyPress(Key.Down, RawInputModifiers.Shift, PhysicalKey.ArrowDown, null);
            Assert.Equal(new CellRange(1, 0, 2, 3), grid.SelectionRange);
            await grid.MergeSelectionAsync(CellMergeMode.Center);
            Assert.Equal(new MergeRange(1, 0, 2, 3), Assert.Single(editor.ViewModel.Merges));
            grid.BorderColor = Colors.Red;
            grid.ApplySelectionBorders(CellBorderPreset.Outside);
            editor.IsReadOnly = true;
            Assert.False(editor.FindControl<Button>("MergeCellsButton")!.IsEnabled);
            Assert.False(editor.FindControl<Button>("BordersButton")!.IsEnabled);
            grid.UnmergeSelection();
            Assert.Single(editor.ViewModel.Merges);
            editor.IsReadOnly = false;
            grid.UnmergeSelection();
            Assert.Empty(editor.ViewModel.Merges);
        }
        finally { window.Close(); }
    }

    [Theory]
    [InlineData("cs")]
    [InlineData("de")]
    [InlineData("es")]
    [InlineData("fr")]
    [InlineData("hu")]
    [InlineData("it")]
    [InlineData("ja")]
    [InlineData("nb")]
    [InlineData("nl")]
    [InlineData("pl")]
    [InlineData("pt")]
    [InlineData("ru")]
    [InlineData("sv")]
    [InlineData("tr")]
    [InlineData("uk")]
    [InlineData("zh-Hans")]
    [InlineData("zh-Hant-TW")]
    public void FormattingCommandsHaveTranslations(string culture)
    {
        CultureInfo original = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo(culture);
            Assert.NotEqual("Merge Cells", EditorText.Get(EditorString.MergeCells));
            Assert.NotEqual("Borders", EditorText.Get(EditorString.Borders));
        }
        finally { CultureInfo.CurrentUICulture = original; }
    }
}
