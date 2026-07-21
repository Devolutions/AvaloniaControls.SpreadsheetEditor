namespace SampleApp;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Avalonia.Media;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using SpreadsheetEditor.Controls.Models;

// App-side .xlsx open/save for the sample host, backed by ClosedXML. Round-trips cell text plus
// whole-cell formatting (bold/italic/underline, font color, fill color). Reads the first
// worksheet and writes a single "Sheet1". Lives in the app, like the TSV/CSV I/O.
public static class XlsxFile
{
    // Reads the first worksheet of an .xlsx stream into a SheetData
    // (cells + column widths + merged ranges), resolving theme colors where possible.
    public static SheetData Read(Stream stream)
    {
        IReadOnlyList<(string Name, SheetData Data)> sheets = ReadWorkbook(stream);
        return sheets.Count > 0
            ? sheets[0].Data
            : new SheetData { Cells = new List<IReadOnlyList<CellData>> { new List<CellData> { CellData.Plain(string.Empty) } } };
    }

    // Reads every worksheet of an .xlsx stream, each as a named SheetData, in
    // workbook (tab) order — so the editor can show one tab per sheet.
    public static IReadOnlyList<(string Name, SheetData Data)> ReadWorkbook(Stream stream)
    {
        // Buffer the bytes so we can read it twice: once via ClosedXML (cells/formatting) and once as
        // a raw zip to extract charts (ClosedXML doesn't expose charts).
        using MemoryStream ms = new();
        stream.CopyTo(ms);
        byte[] bytes = ms.ToArray();

        using XLWorkbook workbook = new(new MemoryStream(bytes));
        List<(string, SheetData)> result = new();
        int index = 0;
        foreach (IXLWorksheet sheet in workbook.Worksheets)
        {
            result.Add((sheet.Name, ReadSheet(workbook, sheet, bytes, index)));
            index++;
        }

        return result;
    }

    // Reads one worksheet into a SheetData. sheetIndex is the
    // 0-based position in the workbook (used to find this sheet's drawing/charts).
    private static SheetData ReadSheet(XLWorkbook workbook, IXLWorksheet sheet, byte[] bytes, int sheetIndex)
    {
        List<List<CellData>> rows = new();
        int lastRow = sheet?.LastRowUsed()?.RowNumber() ?? 0;
        int lastColumn = sheet?.LastColumnUsed()?.ColumnNumber() ?? 0;

        for (int r = 1; r <= lastRow; r++)
        {
            List<CellData> row = new(lastColumn);
            for (int c = 1; c <= lastColumn; c++)
            {
                IXLCell cell = sheet!.Cell(r, c);
                IXLFontBase font = cell.Style.Font;

                // Font color: resolve (incl. theme) but drop default black so plain cells stay unstyled.
                Color? foreground = ResolveColor(workbook, font.FontColor);
                if (foreground is { } f && f.R == 0 && f.G == 0 && f.B == 0)
                {
                    foreground = null;
                }

                Color? background = cell.Style.Fill.PatternType != XLFillPatternValues.None
                    ? ResolveColor(workbook, cell.Style.Fill.BackgroundColor)
                    : null;

                // Import the live formula, or a raw number, so the engine can compute/recompute; the
                // number format renders it (e.g. $50,000). Text/dates keep their formatted string.
                string? format = cell.Style.NumberFormat.Format;
                format = string.IsNullOrWhiteSpace(format) ? null : format;
                string text;
                string? numberFormat = null;
                if (cell.HasFormula)
                {
                    text = "=" + cell.FormulaA1;
                    numberFormat = format;
                }
                else if (cell.DataType == XLDataType.Number)
                {
                    text = cell.GetDouble().ToString(CultureInfo.InvariantCulture);
                    numberFormat = format;
                }
                else
                {
                    text = cell.GetFormattedString(); // text / dates / booleans: show as Excel formatted them
                }

                row.Add(new CellData(
                    text,
                    font.Bold,
                    font.Italic,
                    font.Underline != XLFontUnderlineValues.None,
                    foreground,
                    background,
                    font.FontSize * 4.0 / 3.0,           // points -> device pixels (96/72)
                    AlignmentOf(cell),
                    numberFormat,
                    string.IsNullOrWhiteSpace(font.FontName) ? null : font.FontName,
                    font.Strikethrough));
            }

            rows.Add(row);
        }

        if (rows.Count == 0)
        {
            rows.Add(new List<CellData> { CellData.Plain(string.Empty) });
        }

        List<double> widths = new();
        for (int c = 1; c <= lastColumn; c++)
        {
            // Excel column width (chars) -> pixels (~7px per char + padding), clamped to sane bounds.
            double w = sheet!.Column(c).Width;
            widths.Add(Math.Clamp(w * 7.0 + 5.0, 24.0, 400.0));
        }

        List<MergeRange> merges = new();
        if (sheet is not null)
        {
            foreach (IXLRange range in sheet.MergedRanges)
            {
                IXLAddress first = range.RangeAddress.FirstAddress;
                IXLAddress last = range.RangeAddress.LastAddress;
                merges.Add(new MergeRange(
                    first.RowNumber - 1,
                    first.ColumnNumber - 1,
                    last.RowNumber - first.RowNumber + 1,
                    last.ColumnNumber - first.ColumnNumber + 1));
            }
        }

        List<SheetImage> images = ReadImages(sheet);
        List<SheetChart> charts = XlsxCharts.Read(bytes, sheet, sheetIndex);
        Dictionary<string, string> names = ReadNames(workbook, sheet);
        List<SheetDataBar> dataBars = ReadDataBars(workbook, sheet);

        return new SheetData
        {
            Cells = rows, ColumnWidths = widths, Merges = merges,
            Images = images, Charts = charts, Names = names, DataBars = dataBars,
        };
    }

    // Reads conditional-formatting data bars (range + min/max + color) from the sheet.
    private static List<SheetDataBar> ReadDataBars(XLWorkbook workbook, IXLWorksheet? sheet)
    {
        List<SheetDataBar> bars = new();
        if (sheet is null)
        {
            return bars;
        }

        foreach (IXLConditionalFormat cf in sheet.ConditionalFormats.Where(static c => c.ConditionalFormatType == XLConditionalFormatType.DataBar))
        {
            try
            {
                IXLRangeAddress? rangeAddr = cf.Ranges.FirstOrDefault()?.RangeAddress;
                if (rangeAddr is null)
                {
                    continue;
                }

                IXLAddress first = rangeAddr.FirstAddress;
                IXLAddress last = rangeAddr.LastAddress;
                CellRange range = new(
                    first.RowNumber - 1, first.ColumnNumber - 1,
                    last.RowNumber - first.RowNumber + 1, last.ColumnNumber - first.ColumnNumber + 1);

                double min = cf.Values.TryGetValue(1, out XLFormula? lo) && double.TryParse(lo.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double mn) ? mn : 0;
                double max = cf.Values.TryGetValue(2, out XLFormula? hi) && double.TryParse(hi.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double mx) ? mx : 1;
                Color? color = cf.Colors.TryGetValue(1, out XLColor? col) ? ResolveColor(workbook, col) : null;

                bars.Add(new SheetDataBar
                {
                    Range = range,
                    Min = min,
                    Max = max,
                    Color = color ?? Color.FromRgb(0x63, 0x8E, 0xC6),
                });
            }
            catch
            {
                // skip a data bar we can't read
            }
        }

        return bars;
    }

    // Reads the workbook's defined names, mapping name → its formula text (a range like
    // $O$36 or a formula like SUM('Sheet'!$C$7:$C$10)) so the engine can evaluate them.
    private static Dictionary<string, string> ReadNames(XLWorkbook workbook, IXLWorksheet? sheet)
    {
        Dictionary<string, string> names = new(StringComparer.OrdinalIgnoreCase);
        if (sheet is null)
        {
            return names;
        }

        foreach (IXLDefinedName defined in workbook.DefinedNames)
        {
            string? refersTo = defined.RefersTo?.TrimStart('=');
            if (!string.IsNullOrWhiteSpace(refersTo))
            {
                names[defined.Name] = refersTo;
            }
        }

        return names;
    }

    // Reads embedded pictures, anchored to a cell + pixel offset (free-floating pictures
    // are normalized to an A1 anchor with the offset carrying the absolute position).
    private static List<SheetImage> ReadImages(IXLWorksheet? sheet)
    {
        List<SheetImage> images = new();
        if (sheet is null)
        {
            return images;
        }

        foreach (IXLPicture picture in sheet.Pictures)
        {
            byte[] data;
            try
            {
                Stream stream = picture.ImageStream;
                stream.Position = 0;
                using MemoryStream buffer = new();
                stream.CopyTo(buffer);
                data = buffer.ToArray();
            }
            catch
            {
                continue; // skip a picture we can't read rather than failing the whole open
            }

            if (data.Length == 0)
            {
                continue;
            }

            int row = 0, column = 0, offsetX, offsetY;
            if (picture.Placement == XLPicturePlacement.FreeFloating)
            {
                offsetX = picture.Left;
                offsetY = picture.Top;
            }
            else
            {
                IXLCell anchor = picture.TopLeftCell;
                row = anchor.Address.RowNumber - 1;
                column = anchor.Address.ColumnNumber - 1;
                System.Drawing.Point offset = picture.GetOffset(XLMarkerPosition.TopLeft);
                offsetX = offset.X;
                offsetY = offset.Y;
            }

            images.Add(new SheetImage
            {
                Row = Math.Max(0, row),
                Column = Math.Max(0, column),
                OffsetX = offsetX,
                OffsetY = offsetY,
                Width = picture.Width,
                Height = picture.Height,
                Data = data,
            });
        }

        return images;
    }

    // Resolves an XLColor to an Avalonia color, including theme colors (with tint).
    private static Color? ResolveColor(XLWorkbook workbook, XLColor color)
    {
        try
        {
            switch (color.ColorType)
            {
                case XLColorType.Color:
                {
                    System.Drawing.Color c = color.Color;
                    return Color.FromArgb(c.A, c.R, c.G, c.B);
                }
                case XLColorType.Theme:
                {
                    XLColor? baseColor = ThemeBase(workbook, color.ThemeColor);
                    if (baseColor is null)
                    {
                        return null;
                    }

                    System.Drawing.Color c = baseColor.Color;
                    return ApplyTint(Color.FromArgb(c.A, c.R, c.G, c.B), color.ThemeTint);
                }
                default:
                    return null;
            }
        }
        catch
        {
            return null;
        }
    }

    private static XLColor? ThemeBase(XLWorkbook wb, XLThemeColor theme) => theme switch
    {
        XLThemeColor.Background1 => wb.Theme.Background1,
        XLThemeColor.Text1 => wb.Theme.Text1,
        XLThemeColor.Background2 => wb.Theme.Background2,
        XLThemeColor.Text2 => wb.Theme.Text2,
        XLThemeColor.Accent1 => wb.Theme.Accent1,
        XLThemeColor.Accent2 => wb.Theme.Accent2,
        XLThemeColor.Accent3 => wb.Theme.Accent3,
        XLThemeColor.Accent4 => wb.Theme.Accent4,
        XLThemeColor.Accent5 => wb.Theme.Accent5,
        XLThemeColor.Accent6 => wb.Theme.Accent6,
        _ => null,
    };

    // Applies an Excel tint (-1..1) to a color: negative darkens, positive lightens.
    private static Color ApplyTint(Color c, double tint)
    {
        if (tint == 0)
        {
            return c;
        }

        if (tint < 0)
        {
            double f = 1.0 + tint;
            return Color.FromArgb(c.A, (byte)(c.R * f), (byte)(c.G * f), (byte)(c.B * f));
        }

        return Color.FromArgb(c.A,
            (byte)(c.R + (255 - c.R) * tint),
            (byte)(c.G + (255 - c.G) * tint),
            (byte)(c.B + (255 - c.B) * tint));
    }

    // Writes rows of CellData to a new single-sheet .xlsx workbook with styles.
    public static void Write(Stream stream, IReadOnlyList<IReadOnlyList<CellData>> rows)
    {
        using XLWorkbook workbook = new();
        IXLWorksheet sheet = workbook.Worksheets.Add("Sheet1");

        for (int r = 0; r < rows.Count; r++)
        {
            for (int c = 0; c < rows[r].Count; c++)
            {
                CellData data = rows[r][c];
                IXLCell cell = sheet.Cell(r + 1, c + 1);

                if (data.Text.StartsWith('=') && data.Text.Length > 1)
                {
                    cell.FormulaA1 = data.Text[1..];          // round-trip formulas
                }
                else if (double.TryParse(data.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
                {
                    cell.Value = num;                          // numbers as numbers (so they stay computable)
                }
                else if (!string.IsNullOrEmpty(data.Text))
                {
                    cell.Value = data.Text;
                }

                if (data.NumberFormat is { } nf)
                {
                    cell.Style.NumberFormat.Format = nf;
                }

                if (data.FontFamily is { } fam)
                {
                    cell.Style.Font.FontName = fam;
                }

                if (data.Bold)
                {
                    cell.Style.Font.Bold = true;
                }

                if (data.Italic)
                {
                    cell.Style.Font.Italic = true;
                }

                if (data.Underline)
                {
                    cell.Style.Font.Underline = XLFontUnderlineValues.Single;
                }

                if (data.Strikethrough)
                {
                    cell.Style.Font.Strikethrough = true;
                }

                if (data.Foreground is { } fg)
                {
                    cell.Style.Font.FontColor = ToXl(fg);
                }

                if (data.Background is { } bg)
                {
                    cell.Style.Fill.PatternType = XLFillPatternValues.Solid;
                    cell.Style.Fill.BackgroundColor = ToXl(bg);
                }

                if (data.FontSize > 0)
                {
                    cell.Style.Font.FontSize = data.FontSize * 3.0 / 4.0; // px -> points
                }

                cell.Style.Alignment.Horizontal = data.Alignment switch
                {
                    TextAlignment.Center => XLAlignmentHorizontalValues.Center,
                    TextAlignment.Right => XLAlignmentHorizontalValues.Right,
                    _ => XLAlignmentHorizontalValues.Left,
                };
            }
        }

        workbook.SaveAs(stream);
    }

    // Maps a cell's horizontal alignment, defaulting General to Excel's behavior
    // (numbers right-aligned, text left-aligned).
    private static TextAlignment AlignmentOf(IXLCell cell)
    {
        return cell.Style.Alignment.Horizontal switch
        {
            XLAlignmentHorizontalValues.Center or XLAlignmentHorizontalValues.CenterContinuous => TextAlignment.Center,
            XLAlignmentHorizontalValues.Right => TextAlignment.Right,
            XLAlignmentHorizontalValues.Left => TextAlignment.Left,
            _ => cell.DataType == XLDataType.Number ? TextAlignment.Right : TextAlignment.Left,
        };
    }

    private static XLColor ToXl(Color color) => XLColor.FromArgb(color.A, color.R, color.G, color.B);
}
