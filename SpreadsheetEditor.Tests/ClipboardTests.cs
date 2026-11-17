using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.ViewModels;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using ClosedXML.Excel;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class ClipboardTests
{
    private static readonly DataFormat<byte[]> HtmlFormat = DataFormat.CreateBytesPlatformFormat(
        OperatingSystem.IsWindows() ? "HTML Format" : OperatingSystem.IsMacOS() ? "public.html" : "text/html");

    [AvaloniaFact]
    public async Task PastingAnExcelSizedBlockRebuildsTheViewportOnlyOnce()
    {
        EditorControl editor = new();
        editor.NewSheet(35, 26);
        Window window = new() { Content = editor, Width = 1250, Height = 650 };
        window.Show();
        try
        {
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            grid.SelectCell(new CellAddress(9, 3));
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
            window.UpdateLayout();
            Canvas cells = grid.GetVisualDescendants().OfType<Canvas>()
                .Single(canvas => canvas.Children.OfType<Border>().Any(border => border.Tag is CellAddress));
            int resets = 0;
            cells.Children.CollectionChanged += (_, args) =>
            {
                if (args.Action == NotifyCollectionChangedAction.Remove &&
                    args.OldItems!.OfType<Border>().Any(border => border.Tag is CellAddress(0, 0))) resets++;
            };
            string text = string.Join('\n', Enumerable.Range(0, 3).Select(row =>
                string.Join('\t', Enumerable.Range(0, 13).Select(column => $"Long Excel cell value {row}-{column}"))));
            await window.Clipboard!.SetTextAsync(text);
            await grid.PasteSelectionAsync();
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
            Assert.Equal("Long Excel cell value 2-12", editor.ViewModel.GetCell(new CellAddress(11, 15))!.Text);
            Assert.Equal(1, resets);
        }
        finally
        {
            window.Close();
        }
    }

    [AvaloniaFact]
    public async Task ExcelClassStylesAndValuesArePreservedAndSaved()
    {
        await WithEditor(async (editor, window, grid) =>
        {
            grid.SelectCell(new CellAddress(9, 3));
            Canvas canvas = grid.GetVisualDescendants().OfType<Canvas>()
                .Single(canvas => canvas.Children.OfType<Border>().Any(border => border.Tag is CellAddress));
            int repaints = 0;
            canvas.Children.CollectionChanged += (_, args) =>
            {
                if (args.Action == NotifyCollectionChangedAction.Remove &&
                    args.OldItems!.OfType<Border>().Any(border => border.Tag is CellAddress(0, 0))) repaints++;
            };
            string cells = string.Concat(Enumerable.Range(0, 13).Select(column =>
                $"<td class=xl65 x:num=0.485>{column + 1}</td>"));
            string html = """
                <html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><style><!--
                td {font-family:Calibri; font-size:11pt; vertical-align:bottom;}
                .xl65 {font-weight:700; font-style:italic; color:#FFFF00; background:#123456;
                    text-align:center; vertical-align:middle; white-space:normal;
                    text-decoration:underline; text-underline-style:double;
                    border:.5pt solid #AABBCC; mso-number-format:"0\.0\%";}
                --></style></head><body><!--StartFragment--><table>
                """ + string.Concat(Enumerable.Repeat($"<tr>{cells}</tr>", 3)) + "</table><!--EndFragment--></body></html>";
            // Native Excel clipboard data has a CF_HTML header, UTF-8 bytes and a trailing NUL.
            await SetHtml(window, "Version:1.0\r\nStartHTML:-1\r\nEndHTML:-1\r\n" + html + '\0', "unformatted fallback");
            await grid.PasteSelectionAsync();
            Assert.Equal(1, repaints);
            CellViewModel cell = editor.ViewModel.GetCell(new CellAddress(11, 15))!;
            Assert.Equal("0.485", cell.Text);
            Assert.Equal("48.5%", cell.DisplayText);
            Assert.True(cell.Bold);
            Assert.True(cell.Italic);
            Assert.True(cell.Underline);
            Assert.True(cell.WrapText);
            Assert.Equal(Colors.Yellow, cell.ForegroundColor);
            Assert.Equal(Color.Parse("#123456"), cell.BackgroundColor);
            Assert.Equal("Calibri", cell.FontFamily);
            Assert.Equal(11 * 4.0 / 3.0, cell.FontSize, 5);
            Assert.Equal(TextAlignment.Center, cell.TextAlignment);
            Assert.Equal(VerticalAlignment.Center, cell.Appearance!.VerticalAlignment);
            Assert.True(cell.Appearance.DoubleUnderline);
            Assert.Equal(new CellEdge("thin", Color.Parse("#AABBCC")), cell.Appearance.Left);
            using MemoryStream saved = new();
            await editor.SaveXlsxAsync(saved);
            saved.Position = 0;
            using XLWorkbook workbook = new(saved);
            IXLCell persisted = workbook.Worksheet(1).Cell("P12");
            Assert.Equal(0.485, persisted.GetDouble());
            Assert.True(persisted.Style.Font.Bold);
            Assert.Equal("0.0%", persisted.Style.NumberFormat.Format);
            Assert.Equal(XLColor.FromHtml("#123456"), persisted.Style.Fill.BackgroundColor);
        });
    }

    [AvaloniaFact]
    public async Task HtmlInlineFormattingUnicodeBreaksAndSpansPasteWithOneUndoPoint()
    {
        await WithEditor(async (editor, window, grid) =>
        {
            editor.ViewModel.SetCell(new CellAddress(2, 2), "Previous");
            grid.SelectCell(new CellAddress(2, 2));
            await SetHtml(window, """
                <table><tr><td rowspan=2 style="background-color:#00FF00"><b>Caf&eacute; 🌈<br>Second line</b></td>
                <td colspan=2><font color="#FF0000" face="Arial"><i>Next</i></font></td></tr>
                <tr><td>Bottom</td><td>Right</td></tr></table>
                """);
            await grid.PasteSelectionAsync();
            Assert.Equal("Café 🌈\nSecond line", editor.ViewModel.GetCell(new CellAddress(2, 2))!.Text);
            Assert.True(editor.ViewModel.GetCell(new CellAddress(2, 2))!.Bold);
            Assert.Equal(Colors.Lime, editor.ViewModel.GetCell(new CellAddress(2, 2))!.BackgroundColor);
            Assert.Equal(Colors.Red, editor.ViewModel.GetCell(new CellAddress(2, 3))!.ForegroundColor);
            Assert.True(editor.ViewModel.GetCell(new CellAddress(2, 3))!.Italic);
            Assert.Equal("Right", editor.ViewModel.GetCell(new CellAddress(3, 4))!.Text);
            Assert.Equal(new[] { new MergeRange(2, 2, 2, 1), new MergeRange(2, 3, 1, 2) }, editor.ViewModel.Merges);
            editor.Undo();
            Assert.Equal("Previous", editor.ViewModel.GetCell(new CellAddress(2, 2))!.Text);
            Assert.Empty(editor.ViewModel.Merges);
            Assert.False(editor.ViewModel.CanUndo);
            editor.Redo();
            Assert.Equal(2, editor.ViewModel.Merges.Count);
            Assert.Equal("Right", editor.ViewModel.GetCell(new CellAddress(3, 4))!.Text);
        });
    }

    [AvaloniaTheory]
    [InlineData("<p>Not a table</p>")]
    [InlineData("<table><tr><td colspan=2147483647>Invalid span</td></tr></table>")]
    public async Task UnsupportedHtmlFallsBackToQuotedTsvWithoutSplittingCells(string html)
    {
        await WithEditor(async (editor, window, grid) =>
        {
            await SetHtml(window, html, "\"One\tTwo\"\t\"Three\r\nFour\"\r\n\"Say \"\"hello\"\"\"\tEnd\r\n");
            await grid.PasteSelectionAsync();
            Assert.Equal("One\tTwo", editor.ViewModel.Rows[0].Cells[0].Text);
            Assert.Equal("Three\r\nFour", editor.ViewModel.Rows[0].Cells[1].Text);
            Assert.Equal("Say \"hello\"", editor.ViewModel.Rows[1].Cells[0].Text);
            Assert.Equal("End", editor.ViewModel.Rows[1].Cells[1].Text);
        });
    }

    [AvaloniaFact]
    public async Task InternalCopyKeepsFormulasButExternalHtmlWithTheSameTextUsesItsOwnFormatting()
    {
        await WithEditor(async (editor, window, grid) =>
        {
            CellViewModel source = editor.ViewModel.Rows[0].Cells[0];
            source.ApplyData(CellData.Plain("=1+2") with { Bold = true, Background = Colors.Red });
            grid.CopySelection();
            grid.SelectCell(new CellAddress(1, 0));
            await grid.PasteSelectionAsync();
            Assert.Equal("=1+2", editor.ViewModel.Rows[1].Cells[0].Text);
            Assert.Equal(Colors.Red, editor.ViewModel.Rows[1].Cells[0].BackgroundColor);
            await SetHtml(window, "<table><tr><td style='background-color:blue'>3</td></tr></table>", "3\n");
            grid.SelectCell(new CellAddress(2, 0));
            await grid.PasteSelectionAsync();
            Assert.Equal("3", editor.ViewModel.Rows[2].Cells[0].Text);
            Assert.Equal(Colors.Blue, editor.ViewModel.Rows[2].Cells[0].BackgroundColor);
            Assert.False(editor.ViewModel.Rows[2].Cells[0].Bold);
        });
    }

    [AvaloniaFact]
    public async Task PasteRespectsReadOnlyAndProtectedCells()
    {
        await WithEditor(async (editor, window, grid) =>
        {
            await SetHtml(window, "<table><tr><td style='background:red'>Changed</td></tr></table>");
            editor.IsReadOnly = true;
            await grid.PasteSelectionAsync();
            Assert.Equal(string.Empty, editor.ViewModel.Rows[0].Cells[0].Text);
            editor.IsReadOnly = false;
            editor.ViewModel.ProtectionXml = "<sheetProtection sheet='1'/>";
            await grid.PasteSelectionAsync();
            Assert.Equal(string.Empty, editor.ViewModel.Rows[0].Cells[0].Text);
            CellViewModel cell = editor.ViewModel.Rows[0].Cells[0];
            cell.Appearance = new CellAppearance { Locked = false };
            cell.BackgroundColor = Colors.Yellow;
            await grid.PasteSelectionAsync();
            Assert.Equal("Changed", cell.Text);
            Assert.Equal(Colors.Yellow, cell.BackgroundColor);
            Assert.False(cell.Appearance.Locked);
        });
    }

    [AvaloniaFact]
    public async Task PasteGrowsTheSheetAndAnEmptyClipboardDoesNotReuseTheLastCopy()
    {
        await WithEditor(async (editor, window, grid) =>
        {
            int row = editor.ViewModel.RowCount - 1;
            int column = editor.ViewModel.ColumnCount - 1;
            grid.SelectCell(new CellAddress(row, column));
            await window.Clipboard!.SetTextAsync("One\tTwo\nThree\tFour");
            await grid.PasteSelectionAsync();
            Assert.Equal("Four", editor.ViewModel.GetCell(new CellAddress(row + 1, column + 1))!.Text);
            grid.CopySelection();
            grid.SelectCell(new CellAddress(0, 0));
            await window.Clipboard!.ClearAsync();
            await grid.PasteSelectionAsync();
            Assert.Equal(string.Empty, editor.ViewModel.Rows[0].Cells[0].Text);
        });
    }

    [AvaloniaFact]
    public async Task WaitingForClipboardKeepsDispatcherResponsiveAndDoesNotPasteIntoAnotherSheet()
    {
        await WithEditor(async (editor, window, grid) =>
        {
            DelayedClipboard transfer = new();
            await window.Clipboard!.SetDataAsync(transfer);
            SpreadsheetViewModel original = editor.ViewModel;
            Task paste = grid.PasteSelectionAsync();
            await transfer.Requested.Task;
            bool dispatched = false;
            await Dispatcher.UIThread.InvokeAsync(() => dispatched = true);
            Assert.True(dispatched);
            Assert.False(paste.IsCompleted);
            editor.AddSheet();
            transfer.Value.SetResult(Encoding.UTF8.GetBytes("<table><tr><td>Late data</td></tr></table>"));
            await paste;
            Assert.Equal(string.Empty, editor.ViewModel.Rows[0].Cells[0].Text);
            Assert.Equal(string.Empty, original.Rows[0].Cells[0].Text);
        });
    }

    private static async Task SetHtml(Window window, string html, string text = "fallback")
    {
        DataTransferItem item = new();
        item.SetText(text);
        item.Set(HtmlFormat, Encoding.UTF8.GetBytes(html));
        DataTransfer transfer = new();
        transfer.Add(item);
        await window.Clipboard!.SetDataAsync(transfer);
    }

    private static async Task WithEditor(Func<EditorControl, Window, SpreadsheetGrid, Task> test)
    {
        EditorControl editor = new();
        editor.NewSheet(35, 26);
        Window window = new() { Content = editor, Width = 1000, Height = 550 };
        window.Show();
        try
        {
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
            window.UpdateLayout();
            await test(editor, window, editor.FindControl<SpreadsheetGrid>("Grid")!);
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
        }
    }

    private sealed class DelayedClipboard : IAsyncDataTransfer, IAsyncDataTransferItem
    {
        public TaskCompletionSource Requested { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public TaskCompletionSource<object?> Value { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);
        public IReadOnlyList<DataFormat> Formats => new[] { HtmlFormat };
        public IReadOnlyList<IAsyncDataTransferItem> Items => new[] { this };
        public Task<object?> TryGetRawAsync(DataFormat format)
        {
            Requested.TrySetResult();
            return Value.Task;
        }
        public void Dispose() { }
    }
}
