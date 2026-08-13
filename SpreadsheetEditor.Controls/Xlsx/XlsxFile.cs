namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Avalonia.Media;
using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;
using global::SpreadsheetEditor.Controls.Models;

// .xlsx open/save backed by ClosedXML (+ raw OOXML parsing for charts, via XlsxCharts). Round-trips cell
// text/formulas plus whole-cell formatting (bold/italic/underline/strikethrough, colors, number format,
// font). ReadWorkbook reads every worksheet; Write emits a single "Sheet1". Part of the control library.
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

        // Excel Table definitions, so structured references (Income[[#This Row],[Jan]:[Dec]], etc.) in
        // formulas can be translated to plain A1 ranges the engine understands.
        List<StructuredReferences.TableInfo> tables = StructuredReferences.ReadTables(sheet);

        // Excel Table banding (header / total / striped-row fills) comes from the table style, not from
        // per-cell fills — resolve it so those highlighted rows show up.
        Dictionary<(int Row, int Column), Color> tableFills = BuildTableFills(bytes, sheet, workbook);

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

                // Fall back to the table's banded/header/total fill when the cell has no explicit fill.
                if (background is null && tableFills.TryGetValue((r, c), out Color banded))
                {
                    background = banded;
                }

                // Import the live formula, or a raw number, so the engine can compute/recompute; the
                // number format renders it (e.g. $50,000). Text/dates keep their formatted string.
                // Prefer the explicit format code; for built-in formats (percent, thousands, …) ClosedXML
                // leaves Format empty and only sets NumberFormatId, so map that to its implied code.
                string? format = cell.Style.NumberFormat.Format;
                format = string.IsNullOrWhiteSpace(format)
                    ? BuiltInNumberFormat(cell.Style.NumberFormat.NumberFormatId)
                    : format;
                string text;
                string? numberFormat = null;
                if (cell.HasFormula)
                {
                    text = "=" + StructuredReferences.Translate(cell.FormulaA1, r, c, tables);
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

    // ----- Excel Table banding (header / total / striped-row fills from the table style) -----

    private static readonly XNamespace Ssml = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // Builds a per-cell fill map for the sheet's tables from their table styles (which color the header,
    // totals, and alternating "striped" rows — Excel doesn't store these as per-cell fills).
    private static Dictionary<(int Row, int Column), Color> BuildTableFills(byte[] bytes, IXLWorksheet? sheet, XLWorkbook workbook)
    {
        Dictionary<(int, int), Color> fills = new();
        if (sheet is null || !sheet.Tables.Any())
        {
            return fills;
        }

        try
        {
            // Resolve the workbook's table-style definitions (style name → its element dxf colors).
            using ZipArchive archive = new(new MemoryStream(bytes), ZipArchiveMode.Read);
            ZipArchiveEntry? stylesEntry = archive.GetEntry("xl/styles.xml");
            if (stylesEntry is null)
            {
                return fills;
            }

            XDocument styles = XDocument.Load(stylesEntry.Open());
            List<XElement> dxfs = styles.Descendants(Ssml + "dxfs").Elements(Ssml + "dxf").ToList();

            Dictionary<string, TableStyle> tableStyles = new(StringComparer.OrdinalIgnoreCase);
            foreach (XElement ts in styles.Descendants(Ssml + "tableStyle"))
            {
                string? name = ts.Attribute("name")?.Value;
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                Color? Element(string type, out int size)
                {
                    size = 1;
                    XElement? el = ts.Elements(Ssml + "tableStyleElement").FirstOrDefault(e => e.Attribute("type")?.Value == type);
                    if (el is null || !int.TryParse(el.Attribute("dxfId")?.Value, out int id) || id < 0 || id >= dxfs.Count)
                    {
                        return null;
                    }

                    if (int.TryParse(el.Attribute("size")?.Value, out int s) && s > 0)
                    {
                        size = s;
                    }

                    return DxfFill(dxfs[id], workbook);
                }

                Color? header = Element("headerRow", out _);
                Color? total = Element("totalRow", out _);
                Color? stripe1 = Element("firstRowStripe", out int size1);
                Color? stripe2 = Element("secondRowStripe", out int size2);
                tableStyles[name] = new TableStyle(header, total, stripe1, stripe2, size1, size2);
            }

            foreach (IXLTable table in sheet.Tables)
            {
                string? styleName = table.Theme?.Name;
                if (styleName is null || !tableStyles.TryGetValue(styleName, out TableStyle style))
                {
                    continue;
                }

                IXLRangeAddress range = table.RangeAddress;
                int firstColumn = range.FirstAddress.ColumnNumber;
                int lastColumn = range.LastAddress.ColumnNumber;
                int headerRow = table.ShowHeaderRow ? table.HeadersRow().RowNumber() : -1;
                int totalsRow = table.ShowTotalsRow ? table.TotalsRow().RowNumber() : -1;
                int firstData = table.DataRange?.FirstRow().RowNumber() ?? range.FirstAddress.RowNumber;
                int lastData = table.DataRange?.LastRow().RowNumber() ?? range.LastAddress.RowNumber;

                void Paint(int row, Color? color)
                {
                    if (color is not { } c)
                    {
                        return;
                    }

                    for (int col = firstColumn; col <= lastColumn; col++)
                    {
                        fills[(row, col)] = c;
                    }
                }

                Paint(headerRow, style.Header);
                Paint(totalsRow, style.Total);

                int period = Math.Max(1, style.Stripe1Size) + Math.Max(1, style.Stripe2Size);
                for (int row = firstData; row <= lastData; row++)
                {
                    if (table.ShowRowStripes)
                    {
                        int pos = (row - firstData) % period;
                        Paint(row, pos < Math.Max(1, style.Stripe1Size) ? style.Stripe1 : style.Stripe2);
                    }
                }
            }
        }
        catch
        {
            // Banding is cosmetic — never let it break opening the file.
        }

        return fills;
    }

    private readonly record struct TableStyle(Color? Header, Color? Total, Color? Stripe1, Color? Stripe2, int Stripe1Size, int Stripe2Size);

    // Resolves a dxf's background fill color (rgb, or theme + tint) to an Avalonia color.
    private static Color? DxfFill(XElement dxf, XLWorkbook workbook)
    {
        XElement? bg = dxf.Element(Ssml + "fill")?.Element(Ssml + "patternFill")?.Element(Ssml + "bgColor");
        if (bg is null)
        {
            return null;
        }

        string? rgb = bg.Attribute("rgb")?.Value;
        if (!string.IsNullOrEmpty(rgb) && rgb.Length >= 6)
        {
            string hex = rgb.Length == 8 ? rgb : "FF" + rgb;
            return Color.FromArgb(
                Convert.ToByte(hex[..2], 16), Convert.ToByte(hex.Substring(2, 2), 16),
                Convert.ToByte(hex.Substring(4, 2), 16), Convert.ToByte(hex.Substring(6, 2), 16));
        }

        if (int.TryParse(bg.Attribute("theme")?.Value, out int themeIndex))
        {
            double tint = double.TryParse(bg.Attribute("tint")?.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double t) ? t : 0;
            XLColor? baseColor = ThemeBase(workbook, DxfThemeColor(themeIndex));
            if (baseColor is null)
            {
                return null;
            }

            System.Drawing.Color c = baseColor.Color;
            return ApplyTint(Color.FromArgb(c.A, c.R, c.G, c.B), tint);
        }

        return null;
    }

    // The OOXML dxf/cell theme-color index → the workbook theme slot.
    private static XLThemeColor DxfThemeColor(int index) => index switch
    {
        0 => XLThemeColor.Background1,
        1 => XLThemeColor.Text1,
        2 => XLThemeColor.Background2,
        3 => XLThemeColor.Text2,
        4 => XLThemeColor.Accent1,
        5 => XLThemeColor.Accent2,
        6 => XLThemeColor.Accent3,
        7 => XLThemeColor.Accent4,
        8 => XLThemeColor.Accent5,
        _ => XLThemeColor.Accent6,
    };

    // Writes rows of CellData to a new single-sheet .xlsx workbook with styles.
    public static void Write(
        Stream stream,
        IReadOnlyList<IReadOnlyList<CellData>> rows,
        IReadOnlyDictionary<string, string>? names = null,
        IReadOnlyList<MergeRange>? merges = null,
        string? sheetName = null,
        IReadOnlyList<SheetDataBar>? dataBars = null,
        IReadOnlyList<ChartWriteData>? charts = null,
        IReadOnlyList<double>? columnWidths = null)
    {
        using XLWorkbook workbook = new();
        // Name the worksheet after the sheet so defined names qualified with that sheet (and the
        // engine's single-sheet resolution) still line up on the next open.
        IXLWorksheet sheet = workbook.Worksheets.Add(SafeSheetName(sheetName));

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

        // Persist column widths (px) so the layout survives the round-trip — otherwise every column
        // falls back to Excel's default and the content truncates. Invert the reader's chars->px
        // conversion (px = chars * 7 + 5).
        if (columnWidths is not null)
        {
            for (int c = 0; c < columnWidths.Count; c++)
            {
                double px = columnWidths[c];
                if (px <= 0)
                {
                    continue;
                }

                try
                {
                    sheet.Column(c + 1).Width = Math.Clamp((px - 5.0) / 7.0, 0, 255);
                }
                catch
                {
                    // Skip a column we can't size rather than failing the whole save.
                }
            }
        }

        // Persist merged regions so the layout survives the round-trip.
        if (merges is not null)
        {
            foreach (MergeRange merge in merges)
            {
                if (merge.RowSpan <= 0 || merge.ColumnSpan <= 0 || (merge.RowSpan == 1 && merge.ColumnSpan == 1))
                {
                    continue;
                }

                try
                {
                    sheet.Range(merge.Row + 1, merge.Column + 1, merge.Row + merge.RowSpan, merge.Column + merge.ColumnSpan).Merge();
                }
                catch
                {
                    // Skip a merge we can't recreate rather than failing the whole save.
                }
            }
        }

        // Persist conditional-formatting data bars (the in-cell proportional bars).
        if (dataBars is not null)
        {
            foreach (SheetDataBar bar in dataBars)
            {
                if (bar.Range.RowSpan <= 0 || bar.Range.ColumnSpan <= 0)
                {
                    continue;
                }

                try
                {
                    IXLRange range = sheet.Range(
                        bar.Range.Row + 1, bar.Range.Column + 1,
                        bar.Range.Row + bar.Range.RowSpan, bar.Range.Column + bar.Range.ColumnSpan);
                    range.AddConditionalFormat()
                        .DataBar(ToXl(bar.Color), !bar.ShowValue)
                        .Minimum(XLCFContentType.Number, bar.Min.ToString(CultureInfo.InvariantCulture))
                        .Maximum(XLCFContentType.Number, bar.Max.ToString(CultureInfo.InvariantCulture));
                }
                catch
                {
                    // Skip a data bar we can't recreate rather than failing the whole save.
                }
            }
        }

        // Persist workbook defined names so formulas that reference a name (=NetIncome, etc.) still
        // resolve on the next open — otherwise they come back as #NAME?. A bare range ("$B$1") isn't a
        // valid defined-name target, so qualify it with the sheet; fall back to the raw value.
        if (names is not null)
        {
            foreach (KeyValuePair<string, string> name in names)
            {
                if (string.IsNullOrWhiteSpace(name.Key) || string.IsNullOrWhiteSpace(name.Value))
                {
                    continue;
                }

                string refersTo = name.Value.Trim();
                string qualified = refersTo.Contains('!') ? refersTo : $"'{sheet.Name}'!{refersTo}";
                _ = TryAddName(workbook, name.Key, qualified) || TryAddName(workbook, name.Key, refersTo);
            }
        }

        // ClosedXML can't author charts, so save to memory and graft the chart parts in as raw OOXML
        // (keeps the library pure-managed / cross-platform). No charts → write straight to the stream.
        if (charts is { Count: > 0 })
        {
            using MemoryStream buffer = new();
            workbook.SaveAs(buffer);
            byte[] withCharts = XlsxChartWriter.Inject(buffer.ToArray(), charts);
            stream.Write(withCharts, 0, withCharts.Length);
        }
        else
        {
            workbook.SaveAs(stream);
        }
    }

    private static bool TryAddName(XLWorkbook workbook, string name, string refersTo)
    {
        try
        {
            workbook.DefinedNames.Add(name, refersTo);
            return true;
        }
        catch
        {
            return false;
        }
    }

    // A worksheet name Excel accepts: non-empty, <= 31 chars, without the reserved characters.
    private static string SafeSheetName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            return "Sheet1";
        }

        string result = name.Trim();
        foreach (char bad in new[] { '\\', '/', '?', '*', '[', ']', ':' })
        {
            result = result.Replace(bad, ' ');
        }

        result = result.Trim();
        if (result.Length == 0)
        {
            return "Sheet1";
        }

        return result.Length > 31 ? result[..31] : result;
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

    // Format code for a built-in Excel numFmtId (0-163) that ClosedXML leaves as an empty Format string.
    // Only the numeric formats the renderer understands are mapped; dates/text fall through to null
    // (General / the cell's formatted string). https://learn.microsoft.com built-in number formats.
    private static string? BuiltInNumberFormat(int numberFormatId) => numberFormatId switch
    {
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        37 or 38 => "#,##0;(#,##0)",
        39 or 40 => "#,##0.00;(#,##0.00)",
        _ => null,
    };
}
