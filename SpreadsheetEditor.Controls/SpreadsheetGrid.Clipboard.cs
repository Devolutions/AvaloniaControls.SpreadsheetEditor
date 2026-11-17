namespace SpreadsheetEditor.Controls;

using System;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;

using global::SpreadsheetEditor.Controls.Clipboard;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.ViewModels;

public partial class SpreadsheetGrid
{
    private static readonly DataFormat<byte[]> HtmlClipboardFormat = DataFormat.CreateBytesPlatformFormat(
        OperatingSystem.IsWindows() ? "HTML Format" : OperatingSystem.IsMacOS() ? "public.html" : "text/html");
    private bool pasting;

    public async void PasteSelection() => await this.PasteSelectionAsync();

    // The awaitable entry point also lets hosts wait for clipboard parsing to finish.
    public async Task PasteSelectionAsync()
    {
        if (this.sheet is not { } sheet || this.ReadOnly || this.pasting) return;
        CellRange range = this.SelectionRange;
        this.pasting = true;
        try
        {
            ClipboardTable? table = await this.ReadClipboardTableAsync();
            if (table is null || table.Cells.Length == 0 || this.ReadOnly || !ReferenceEquals(this.sheet, sheet)) return;
            CellData[][] block = table.Cells;
            int width = block.Max(row => row.Length);
            CellRange targetRange = new(range.Row, range.Column, block.Length, width);
            if (!sheet.CanEdit(targetRange)) return;

            sheet.Checkpoint();
            this.ApplyCellUpdates(() => sheet.RunBatch(() =>
            {
                sheet.EnsureSize(range.Row + block.Length, range.Column + width);
                for (int i = 0; i < block.Length; i++)
                {
                    for (int j = 0; j < block[i].Length; j++)
                    {
                        CellViewModel target = sheet.Rows[range.Row + i].Cells[range.Column + j];
                        if (sheet.IsProtected) target.Text = block[i][j].Text;
                        else target.ApplyData(block[i][j]);
                    }
                }
                if (!sheet.IsProtected && table.Merges.Count > 0)
                {
                    sheet.SetMerges(sheet.Merges.Where(merge => !Intersects(merge, targetRange))
                        .Concat(table.Merges.Select(merge => merge with { Row = range.Row + merge.Row, Column = range.Column + merge.Column })).ToArray());
                }
            }));
            this.Rebuild();
        }
        finally
        {
            this.pasting = false;
        }
    }

    private async Task<ClipboardTable?> ReadClipboardTableAsync()
    {
        IClipboard? clip = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clip is null) return this.clipboard is { } block ? new ClipboardTable(block, Array.Empty<MergeRange>()) : null;
        try
        {
            using IAsyncDataTransfer? data = await clip.TryGetDataAsync();
            if (data is null) return null;
            if (await data.TryGetValueAsync(HtmlClipboardFormat) is { Length: > 0 and <= 16_777_216 } html)
            {
                ClipboardTable? table = await Task.Run(() =>
                {
                    try { return ClipboardTable.FromHtml(Encoding.UTF8.GetString(html)); }
                    catch (Exception ex) when (ex is FormatException or ArgumentException or RegexMatchTimeoutException) { return null; }
                });
                if (table is not null) return table;
            }
            string? text = await data.TryGetTextAsync();
            if (text is not null && text == this.clipboardText && this.clipboard is { } internalBlock)
                return new ClipboardTable(internalBlock, Array.Empty<MergeRange>());
            return await Task.Run(() => ClipboardTable.FromText(text));
        }
        catch
        {
            // Another process can temporarily own or replace the clipboard during a paste.
            return null;
        }
    }

    private static bool Intersects(MergeRange merge, CellRange range) =>
        merge.Row < range.Row + range.RowSpan && merge.Row + merge.RowSpan > range.Row &&
        merge.Column < range.Column + range.ColumnSpan && merge.Column + merge.ColumnSpan > range.Column;
}
