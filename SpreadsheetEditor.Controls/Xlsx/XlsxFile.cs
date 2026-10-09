namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

using Avalonia.Media;

using ClosedXML.Excel;
using ClosedXML.Excel.Drawings;

using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Localization;
using global::SpreadsheetEditor.Controls.Models;

// .xlsx open/save backed by ClosedXML (+ raw OOXML parsing for charts, via XlsxCharts). Round-trips cell
// text/formulas plus whole-cell formatting (bold/italic/underline/strikethrough, colors, number format,
// font). ReadWorkbook and WriteWorkbook preserve every worksheet in tab order.
internal sealed record XlsxSheetPreview(int Index, string Name, string PartPath, Color? TabColor);

internal sealed record XlsxWorkbookPreview(IReadOnlyList<XlsxSheetPreview> Sheets, SheetData FirstSheet);

internal sealed record XlsxSheetWriteData(
    string Name,
    SheetData Data,
    IReadOnlyList<ChartWriteData> Charts);

// Owns one ClosedXML workbook for the lifetime of a staged import. Selected worksheets are converted
// serially so rapid tab switching cannot create several simultaneous full-workbook parsers.
internal sealed class XlsxWorkbookReader(byte[] bytes)
{
    private readonly SemaphoreSlim gate = new(1, 1);
    private XLWorkbook? workbook;
    private volatile bool disposeRequested;

    internal async Task<SheetData> ReadWorksheetAsync(int sheetIndex, CancellationToken cancellationToken)
    {
        await this.gate.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            if (this.disposeRequested)
            {
                throw new ObjectDisposedException(nameof(XlsxWorkbookReader));
            }

            return await Task.Run(
                () =>
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    this.workbook ??= new XLWorkbook(new MemoryStream(bytes, writable: false));
                    cancellationToken.ThrowIfCancellationRequested();
                    IXLWorksheet? sheet = this.workbook.Worksheets.Skip(sheetIndex).FirstOrDefault();
                    if (sheet is null)
                    {
                        throw new ArgumentOutOfRangeException(nameof(sheetIndex));
                    }

                    return XlsxFile.ReadSheet(this.workbook, sheet, bytes, sheetIndex, cancellationToken);
                },
                cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            this.gate.Release();
        }
    }

    internal void DisposeWhenIdle()
    {
        this.disposeRequested = true;
        _ = Task.Run(async () =>
        {
            await this.gate.WaitAsync().ConfigureAwait(false);
            try
            {
                this.workbook?.Dispose();
                this.workbook = null;
            }
            finally
            {
                this.gate.Release();
            }
        });
    }
}

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

        List<(string, SheetData)> result = new();
        ReadWorkbookIncrementally(bytes, (index, name, data) => result.Add((name, data)), CancellationToken.None);

        return result;
    }

    // Reads the workbook once and publishes each completed worksheet immediately. Callers run this
    // on a worker thread and can progressively replace placeholder tabs without retaining every
    // SheetData snapshot at once.
    internal static void ReadWorkbookIncrementally(
        byte[] bytes,
        Action<int, string, SheetData> sheetRead,
        CancellationToken cancellationToken)
    {
        using XLWorkbook workbook = new(new MemoryStream(bytes, writable: false));
        int index = 0;
        foreach (IXLWorksheet sheet in workbook.Worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            SheetData data = ReadSheet(workbook, sheet, bytes, index, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            sheetRead(index, sheet.Name, data);
            index++;
        }
    }

    // Reads one full-fidelity worksheet. ClosedXML still has to open the package, so callers keep
    // this work off the UI thread and invoke it only after the worksheet has been selected.
    internal static SheetData ReadWorksheet(byte[] bytes, int sheetIndex, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using XLWorkbook workbook = new(new MemoryStream(bytes, writable: false));
        cancellationToken.ThrowIfCancellationRequested();
        IXLWorksheet? sheet = workbook.Worksheets.Skip(sheetIndex).FirstOrDefault();
        if (sheet is null)
        {
            throw new ArgumentOutOfRangeException(nameof(sheetIndex));
        }

        SheetData data = ReadSheet(workbook, sheet, bytes, sheetIndex, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return data;
    }

    // Produces a fast, read-only first-sheet preview directly from the OOXML package. This avoids
    // ClosedXML's eager whole-workbook construction, which dominates open time for large files.
    // The normal reader subsequently replaces this preview with the full-fidelity worksheet.
    internal static XlsxWorkbookPreview ReadPreview(
        byte[] bytes,
        CancellationToken cancellationToken,
        int rowLimit = 128,
        int columnLimit = 64)
    {
        using ZipArchive archive = new(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        XDocument workbook = ReadXmlPart(archive, "xl/workbook.xml");
        XDocument relationships = ReadXmlPart(archive, "xl/_rels/workbook.xml.rels");
        XNamespace relationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";

        Dictionary<string, string> targets = relationships
            .Descendants(packageRelationships + "Relationship")
            .Where(static relationship =>
                relationship.Attribute("Id") is not null &&
                relationship.Attribute("Target") is not null &&
                relationship.Attribute("Type")?.Value.EndsWith("/worksheet", StringComparison.Ordinal) == true)
            .ToDictionary(
                relationship => relationship.Attribute("Id")!.Value,
                relationship => ResolvePartPath("xl/workbook.xml", relationship.Attribute("Target")!.Value),
                StringComparer.Ordinal);

        IReadOnlyList<Color?> theme = ReadThemeColors(archive);
        List<XlsxSheetPreview> sheets = new();
        foreach (XElement sheet in workbook.Descendants(Ssml + "sheet"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            string name = sheet.Attribute("name")?.Value?.Trim() ?? string.Empty;
            string relationshipId = sheet.Attribute(relationshipNamespace + "id")?.Value ?? string.Empty;
            if (targets.TryGetValue(relationshipId, out string? part))
            {
                string resolvedName = string.IsNullOrWhiteSpace(name) ? $"{EditorText.Get(EditorString.Sheet)}{sheets.Count + 1}" : name;
                Color? tabColor = ReadPreviewTabColor(archive, part, theme, cancellationToken);
                sheets.Add(new XlsxSheetPreview(sheets.Count, resolvedName, part, tabColor));
            }
        }

        if (sheets.Count == 0)
        {
            return new XlsxWorkbookPreview(
                new[] { new XlsxSheetPreview(0, EditorText.Get(EditorString.Sheet) + "1", string.Empty, null) },
                new SheetData { Cells = new List<IReadOnlyList<CellData>> { new List<CellData> { CellData.Plain(string.Empty) } } });
        }

        cancellationToken.ThrowIfCancellationRequested();
        SheetData firstSheet = ReadPreviewSheet(
            archive,
            bytes,
            sheets[0].PartPath,
            sheets[0].Index,
            sheets[0].TabColor,
            rowLimit,
            columnLimit,
            cancellationToken);
        return new XlsxWorkbookPreview(sheets, firstSheet);
    }

    // Reads only a screen-sized slice of a selected worksheet. The metadata object carries its
    // already-resolved tab color so this pass does not need to rescan any unrelated worksheets.
    internal static SheetData ReadWorksheetPreview(
        byte[] bytes,
        XlsxSheetPreview sheet,
        int rowLimit,
        int columnLimit,
        CancellationToken cancellationToken)
    {
        using ZipArchive archive = new(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        return ReadPreviewSheet(
            archive,
            bytes,
            sheet.PartPath,
            sheet.Index,
            sheet.TabColor,
            rowLimit,
            columnLimit,
            cancellationToken);
    }

    private static SheetData ReadPreviewSheet(
        ZipArchive archive,
        byte[] bytes,
        string partPath,
        int sheetIndex,
        Color? tabColor,
        int rowLimit,
        int columnLimit,
        CancellationToken cancellationToken)
    {
        PreviewWorksheetSlice slice = ReadPreviewWorksheetSlice(
            archive,
            partPath,
            rowLimit,
            columnLimit,
            cancellationToken);
        CellRange? autoFilterRange = ReadAutoFilterRange(archive, partPath, cancellationToken);
        List<XElement> cells = slice.Cells;
        HashSet<int> sharedStringIndexes = new();
        foreach (XElement cell in cells)
        {
            if (cell.Attribute("t")?.Value == "s" &&
                int.TryParse(cell.Element(Ssml + "v")?.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int index))
            {
                sharedStringIndexes.Add(index);
            }
        }

        Dictionary<int, string> sharedStrings = ReadSharedStrings(archive, sharedStringIndexes, cancellationToken);
        List<PreviewCellStyle> styles = ReadPreviewStyles(archive);
        Dictionary<(int Row, int Column), CellData> values = new();
        int lastRow = slice.RowCount;
        int lastColumn = slice.ColumnCount;
        foreach (XElement cell in cells)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryParseCellAddress(cell.Attribute("r")?.Value, out int row, out int column))
            {
                continue;
            }

            int styleIndex = int.TryParse(cell.Attribute("s")?.Value, out int parsedStyle) ? parsedStyle : 0;
            PreviewCellStyle style = styleIndex >= 0 && styleIndex < styles.Count ? styles[styleIndex] : default;
            (string text, string? formulaResult) = PreviewCellText(cell, sharedStrings, style.NumberFormat);
            double? formulaNumber = cell.Element(Ssml + "f") is not null &&
                double.TryParse(cell.Element(Ssml + "v")?.Value, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out double cachedNumber)
                    ? cachedNumber : null;
            values[(row, column)] = new CellData(
                text,
                style.Bold,
                style.Italic,
                style.Underline,
                style.Foreground,
                style.Background,
                style.FontSize,
                style.Alignment,
                style.NumberFormat,
                style.FontFamily,
                style.Strikethrough,
                formulaResult,
                FormulaNumericValue: formulaNumber);
        }

        lastRow = Math.Max(1, lastRow);
        lastColumn = Math.Max(1, lastColumn);
        List<List<CellData>> rows = new(lastRow);
        for (int row = 0; row < lastRow; row++)
        {
            List<CellData> rowData = new(lastColumn);
            for (int column = 0; column < lastColumn; column++)
            {
                rowData.Add(values.TryGetValue((row, column), out CellData value) ? value : CellData.Plain(string.Empty));
            }
            rows.Add(rowData);
        }

        // A huge sheet keeps the first-screen preview fast; its full import applies scales later.
        if ((archive.GetEntry(partPath)?.Length ?? 0) <= 2_000_000)
        {
            ApplyColorScales(rows, ReadColorScales(archive, partPath, ReadThemeColors(archive)), cancellationToken);
        }
        (List<SheetDataBar> previewBars, List<SheetCellValueRule> previewValueRules) =
            ReadConditionalFormatting(archive, partPath, ReadThemeColors(archive), cancellationToken);

        List<double> widths = Enumerable.Repeat(0d, lastColumn).ToList();
        foreach ((int min, int max, double width) in slice.Columns)
        {
            double pixels = Math.Clamp(width * 7.0 + 5.0, 24.0, 400.0);
            for (int index = Math.Max(1, min); index <= Math.Min(lastColumn, max); index++)
            {
                widths[index - 1] = pixels;
            }
        }

        return new SheetData
        {
            Cells = rows,
            ColumnWidths = widths,
            Merges = slice.Merges,
            Charts = XlsxCharts.Read(bytes, sheet: null, sheetIndex),
            DataBars = previewBars,
            CellValueRules = previewValueRules,
            AutoFilterRange = autoFilterRange,
            HiddenRows = slice.HiddenRows,
            TabColor = tabColor ?? PreviewColor(slice.TabColor, ReadThemeColors(archive)),
        };
    }

    private static PreviewWorksheetSlice ReadPreviewWorksheetSlice(
        ZipArchive archive,
        string partPath,
        int rowLimit,
        int columnLimit,
        CancellationToken cancellationToken)
    {
        rowLimit = Math.Clamp(rowLimit, 1, 2048);
        columnLimit = Math.Clamp(columnLimit, 1, 256);
        ZipArchiveEntry entry = archive.GetEntry(partPath) ?? throw new InvalidDataException(EditorText.Format(EditorString.MissingWorkbookPart, partPath));
        List<XElement> cells = new();
        List<(int Min, int Max, double Width)> columns = new();
        List<MergeRange> merges = new();
        List<int> hiddenRows = new();
        XElement? tabColor = null;
        int dimensionRows = 0;
        int dimensionColumns = 0;
        int lastLoadedRow = 0;
        int lastLoadedColumn = 0;

        using Stream stream = entry.Open();
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = true,
        });
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            if (reader.LocalName == "dimension" &&
                CellRange.TryParse(reader.GetAttribute("ref"), out CellRange dimension))
            {
                dimensionRows = dimension.Row + dimension.RowSpan;
                dimensionColumns = dimension.Column + dimension.ColumnSpan;
                reader.Read();
                continue;
            }

            if (reader.LocalName == "sheetPr")
            {
                XElement properties = (XElement)XNode.ReadFrom(reader);
                tabColor = properties.Element(Ssml + "tabColor");
                continue;
            }

            if (reader.LocalName == "col")
            {
                if (int.TryParse(reader.GetAttribute("min"), out int min) &&
                    int.TryParse(reader.GetAttribute("max"), out int max) &&
                    double.TryParse(reader.GetAttribute("width"), NumberStyles.Float, CultureInfo.InvariantCulture, out double width))
                {
                    columns.Add((min, max, width));
                    lastLoadedColumn = Math.Max(lastLoadedColumn, Math.Min(columnLimit, max));
                }
                reader.Read();
                continue;
            }

            if (reader.LocalName == "row")
            {
                int oneBasedRow = int.TryParse(reader.GetAttribute("r"), out int parsedRow)
                    ? parsedRow
                    : lastLoadedRow + 1;
                if (oneBasedRow > rowLimit)
                {
                    break;
                }

                if (reader.GetAttribute("hidden") is "1" or "true")
                {
                    hiddenRows.Add(oneBasedRow - 1);
                }

                XElement row = (XElement)XNode.ReadFrom(reader);
                foreach (XElement cell in row.Descendants(Ssml + "c"))
                {
                    if (!TryParseCellAddress(cell.Attribute("r")?.Value, out _, out int column) || column >= columnLimit)
                    {
                        continue;
                    }
                    cells.Add(cell);
                    lastLoadedColumn = Math.Max(lastLoadedColumn, column + 1);
                }
                lastLoadedRow = Math.Max(lastLoadedRow, oneBasedRow);
                continue;
            }

            // Merge definitions come after sheetData. They are available without scanning a large
            // sheet only when the first slice already contains the full worksheet.
            if (dimensionRows <= rowLimit && reader.LocalName == "mergeCell" &&
                CellRange.TryParse(reader.GetAttribute("ref"), out CellRange range) &&
                range.Row < rowLimit && range.Column < columnLimit)
            {
                int rowSpan = Math.Min(range.RowSpan, rowLimit - range.Row);
                int columnSpan = Math.Min(range.ColumnSpan, columnLimit - range.Column);
                merges.Add(new MergeRange(range.Row, range.Column, rowSpan, columnSpan));
            }
            reader.Read();
        }

        int rowCount = Math.Max(1, Math.Min(rowLimit, dimensionRows > 0 ? dimensionRows : lastLoadedRow));
        int columnCount = Math.Max(1, Math.Min(columnLimit, dimensionColumns > 0 ? dimensionColumns : lastLoadedColumn));
        return new PreviewWorksheetSlice(cells, columns, merges, hiddenRows, tabColor, rowCount, columnCount);
    }

    // AutoFilter follows sheetData in OOXML, so a limited first-screen preview would otherwise
    // miss it. Skip the large cell subtree without allocating it to find the header range.
    private static CellRange? ReadAutoFilterRange(
        ZipArchive archive, string partPath, CancellationToken cancellationToken)
    {
        if (archive.GetEntry(partPath) is not { } entry)
        {
            return null;
        }

        using Stream stream = entry.Open();
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = true,
        });
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheetData")
            {
                reader.Skip();
                continue;
            }
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "autoFilter")
            {
                return CellRange.TryParse(reader.GetAttribute("ref"), out CellRange range) ? range : null;
            }
            reader.Read();
        }
        return null;
    }

    private static Dictionary<int, string> ReadSharedStrings(
        ZipArchive archive,
        HashSet<int> requested,
        CancellationToken cancellationToken)
    {
        Dictionary<int, string> result = new();
        if (requested.Count == 0 || archive.GetEntry("xl/sharedStrings.xml") is not { } entry)
        {
            return result;
        }

        int maximum = requested.Max();
        using Stream stream = entry.Open();
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        int index = 0;
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "si")
            {
                reader.Read();
                continue;
            }

            XElement item = (XElement)XNode.ReadFrom(reader);
            if (requested.Contains(index))
            {
                result[index] = string.Concat(item.Descendants(Ssml + "t").Select(static text => text.Value));
            }
            if (index >= maximum)
            {
                break;
            }
            index++;
        }

        return result;
    }

    private static (string Text, string? FormulaResult) PreviewCellText(
        XElement cell,
        IReadOnlyDictionary<int, string> sharedStrings,
        string? numberFormat)
    {
        XElement? formula = cell.Element(Ssml + "f");
        string type = cell.Attribute("t")?.Value ?? string.Empty;
        string raw = cell.Element(Ssml + "v")?.Value ?? string.Empty;
        string value = type switch
        {
            "s" when int.TryParse(raw, out int index) && sharedStrings.TryGetValue(index, out string? shared) => shared,
            "inlineStr" => string.Concat(cell.Descendants(Ssml + "t").Select(static text => text.Value)),
            "b" => raw == "1" ? "TRUE" : "FALSE",
            _ => raw,
        };

        // Shared-formula followers can carry an empty <f/>. Show their cached value in the small
        // preview; ClosedXML expands the formula when the full sheet is imported.
        string cachedResult = type.Length == 0 &&
            double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double number)
                ? NumberFormatter.Format(number, numberFormat)
                : value;
        return formula is not null && !string.IsNullOrWhiteSpace(formula.Value)
            ? ("=" + formula.Value, cachedResult)
            : (value, null);
    }

    private static List<PreviewCellStyle> ReadPreviewStyles(ZipArchive archive)
    {
        if (archive.GetEntry("xl/styles.xml") is not { } entry)
        {
            return new List<PreviewCellStyle> { default };
        }

        XDocument document;
        using (Stream stream = entry.Open())
        {
            document = XDocument.Load(stream);
        }

        IReadOnlyList<Color?> theme = ReadThemeColors(archive);
        List<XElement> fonts = document.Descendants(Ssml + "fonts").Elements(Ssml + "font").ToList();
        List<XElement> fills = document.Descendants(Ssml + "fills").Elements(Ssml + "fill").ToList();
        Dictionary<int, string> numberFormats = document.Descendants(Ssml + "numFmt")
            .Where(static format => int.TryParse(format.Attribute("numFmtId")?.Value, out _) && format.Attribute("formatCode") is not null)
            .ToDictionary(
                format => int.Parse(format.Attribute("numFmtId")!.Value, CultureInfo.InvariantCulture),
                format => format.Attribute("formatCode")!.Value);

        List<PreviewCellStyle> result = new();
        foreach (XElement xf in document.Descendants(Ssml + "cellXfs").Elements(Ssml + "xf"))
        {
            int fontId = int.TryParse(xf.Attribute("fontId")?.Value, out int parsedFont) ? parsedFont : 0;
            int fillId = int.TryParse(xf.Attribute("fillId")?.Value, out int parsedFill) ? parsedFill : 0;
            int numberFormatId = int.TryParse(xf.Attribute("numFmtId")?.Value, out int parsedFormat) ? parsedFormat : 0;
            XElement? font = fontId >= 0 && fontId < fonts.Count ? fonts[fontId] : null;
            XElement? fill = fillId >= 0 && fillId < fills.Count ? fills[fillId] : null;
            Color? foreground = PreviewColor(font?.Element(Ssml + "color"), theme);
            if (foreground is { R: 0, G: 0, B: 0 })
            {
                foreground = null;
            }

            string? numberFormat = numberFormats.TryGetValue(numberFormatId, out string? custom)
                ? custom
                : BuiltInNumberFormat(numberFormatId);
            string horizontal = xf.Element(Ssml + "alignment")?.Attribute("horizontal")?.Value ?? string.Empty;
            result.Add(new PreviewCellStyle(
                font?.Element(Ssml + "b") is not null,
                font?.Element(Ssml + "i") is not null,
                font?.Element(Ssml + "u") is not null,
                font?.Element(Ssml + "strike") is not null,
                foreground,
                PreviewColor(fill?.Element(Ssml + "patternFill")?.Element(Ssml + "fgColor"), theme),
                double.TryParse(font?.Element(Ssml + "sz")?.Attribute("val")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double size)
                    ? size * 4.0 / 3.0
                    : 0,
                font?.Element(Ssml + "name")?.Attribute("val")?.Value,
                numberFormat,
                horizontal is "center" or "centerContinuous" ? TextAlignment.Center : horizontal == "right" ? TextAlignment.Right : TextAlignment.Left));
        }

        return result.Count == 0 ? new List<PreviewCellStyle> { default } : result;
    }

    private static IReadOnlyList<Color?> ReadThemeColors(ZipArchive archive)
    {
        Color?[] theme = new Color?[12];
        if (archive.GetEntry("xl/theme/theme1.xml") is not { } entry)
        {
            return theme;
        }

        XDocument document;
        using (Stream stream = entry.Open())
        {
            document = XDocument.Load(stream);
        }

        XNamespace drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        XElement? scheme = document.Descendants(drawing + "clrScheme").FirstOrDefault();
        if (scheme is null)
        {
            return theme;
        }

        string[] order = { "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink" };
        for (int index = 0; index < order.Length; index++)
        {
            XElement? color = scheme.Element(drawing + order[index])?.Elements().FirstOrDefault();
            string? hex = color?.Attribute("lastClr")?.Value ?? color?.Attribute("val")?.Value;
            theme[index] = ParseRgb(hex);
        }

        return theme;
    }

    // sheetPr is located before sheetData, so tab color can be read without scanning cell rows.
    // This makes every tab look final during the initial workbook metadata pass.
    private static Color? ReadPreviewTabColor(
        ZipArchive archive,
        string partPath,
        IReadOnlyList<Color?> theme,
        CancellationToken cancellationToken)
    {
        ZipArchiveEntry entry = archive.GetEntry(partPath) ?? throw new InvalidDataException(EditorText.Format(EditorString.MissingWorkbookPart, partPath));
        using Stream stream = entry.Open();
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = true,
        });
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            if (reader.LocalName == "sheetPr")
            {
                XElement properties = (XElement)XNode.ReadFrom(reader);
                return PreviewColor(properties.Element(Ssml + "tabColor"), theme);
            }

            if (reader.LocalName == "sheetData")
            {
                break;
            }

            reader.Read();
        }

        return null;
    }

    private static Color? PreviewColor(XElement? element, IReadOnlyList<Color?> theme)
    {
        if (element is null)
        {
            return null;
        }

        Color? color = ParseRgb(element.Attribute("rgb")?.Value);
        if (color is null && int.TryParse(element.Attribute("theme")?.Value, out int themeIndex) &&
            themeIndex >= 0 && themeIndex < theme.Count)
        {
            color = theme[themeIndex];
        }

        if (color is { } resolved &&
            double.TryParse(element.Attribute("tint")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out double tint))
        {
            color = ApplyTint(resolved, tint);
        }
        return color;
    }

    private static Color? ParseRgb(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        string hex = value.Trim().TrimStart('#');
        if (hex.Length == 8)
        {
            hex = hex[2..];
        }
        return hex.Length == 6 && uint.TryParse(hex, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out uint rgb)
            ? Color.FromRgb((byte)(rgb >> 16), (byte)(rgb >> 8), (byte)rgb)
            : null;
    }

    private static XDocument ReadXmlPart(ZipArchive archive, string path)
    {
        ZipArchiveEntry entry = archive.GetEntry(path) ?? throw new InvalidDataException(EditorText.Format(EditorString.MissingWorkbookPart, path));
        using Stream stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static string ResolvePartPath(string sourcePart, string target)
    {
        Uri packageRoot = new("http://package/");
        Uri source = new(packageRoot, sourcePart);
        return Uri.UnescapeDataString(new Uri(source, target.Replace('\\', '/')).AbsolutePath.TrimStart('/'));
    }

    private static bool TryParseCellAddress(string? reference, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(reference))
        {
            return false;
        }

        string clean = reference.Replace("$", string.Empty, StringComparison.Ordinal);
        int split = 0;
        while (split < clean.Length && char.IsLetter(clean[split]))
        {
            split++;
        }
        if (split == 0 || split == clean.Length || !int.TryParse(clean[split..], out int oneBasedRow))
        {
            return false;
        }

        column = CellAddress.LettersToColumn(clean[..split]);
        row = oneBasedRow - 1;
        return row >= 0 && column >= 0;
    }

    private readonly record struct PreviewCellStyle(
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strikethrough,
        Color? Foreground,
        Color? Background,
        double FontSize,
        string? FontFamily,
        string? NumberFormat,
        TextAlignment Alignment);

    private sealed record PreviewWorksheetSlice(
        List<XElement> Cells,
        List<(int Min, int Max, double Width)> Columns,
        List<MergeRange> Merges,
        List<int> HiddenRows,
        XElement? TabColor,
        int RowCount,
        int ColumnCount);

    // Reads one worksheet into a SheetData. sheetIndex is the
    // 0-based position in the workbook (used to find this sheet's drawing/charts).
    internal static SheetData ReadSheet(
        XLWorkbook workbook,
        IXLWorksheet sheet,
        byte[] bytes,
        int sheetIndex,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        List<List<CellData>> rows = new();
        List<int> hiddenRows = new();
        int lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        int lastColumn = sheet.LastColumnUsed()?.ColumnNumber() ?? 0;
        CellRange? autoFilterRange;
        using (ZipArchive archive = new(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read))
        {
            autoFilterRange = ReadAutoFilterRange(archive, WorksheetPartPath(archive, sheetIndex), cancellationToken);
        }
        if (autoFilterRange is { } filterRange)
        {
            lastRow = Math.Max(lastRow, filterRange.Row + filterRange.RowSpan);
            lastColumn = Math.Max(lastColumn, filterRange.Column + filterRange.ColumnSpan);
        }

        // Empty cells inside an Excel Table are still part of the visible formatted table. ClosedXML's
        // LastRowUsed/LastColumnUsed only report cells with content, so include every table's declared
        // range or blank banded rows disappear after import.
        foreach (IXLTable table in sheet.Tables)
        {
            lastRow = Math.Max(lastRow, table.RangeAddress.LastAddress.RowNumber);
            lastColumn = Math.Max(lastColumn, table.RangeAddress.LastAddress.ColumnNumber);
        }

        // Excel Table definitions, so structured references (Income[[#This Row],[Jan]:[Dec]], etc.) in
        // formulas can be translated to plain A1 ranges the engine understands.
        List<StructuredReferences.TableInfo> tables = StructuredReferences.ReadTables(sheet);

        // Excel Table formatting (header / total / striped rows) comes from the table style, not from
        // per-cell styles. Resolve both custom styles embedded in styles.xml and Excel's built-in
        // TableStyleLight/Medium/Dark families.
        Dictionary<(int Row, int Column), TableCellStyle> tableStyles = BuildTableStyles(bytes, sheet, workbook);
        Dictionary<IXLStyle, ImportedCellStyle> styleCache = new();

        for (int r = 1; r <= lastRow; r++)
        {
            if ((r & 63) == 1)
            {
                cancellationToken.ThrowIfCancellationRequested();
            }

            List<CellData> row = new(lastColumn);
            if (sheet.Row(r).IsHidden)
            {
                hiddenRows.Add(r - 1);
            }
            for (int c = 1; c <= lastColumn; c++)
            {
                IXLCell cell = sheet!.Cell(r, c);
                IXLStyle xlStyle = cell.Style;
                if (!styleCache.TryGetValue(xlStyle, out ImportedCellStyle importedStyle))
                {
                    importedStyle = ReadCellStyle(workbook, xlStyle);
                    styleCache[xlStyle] = importedStyle;
                }

                Color? foreground = importedStyle.Foreground;
                Color? background = importedStyle.Background;
                bool bold = importedStyle.Bold;
                if (tableStyles.TryGetValue((r, c), out TableCellStyle tableStyle))
                {
                    // Direct cell formatting wins; table formatting supplies only missing values.
                    background ??= tableStyle.Background;
                    foreground ??= tableStyle.Foreground;
                    bold |= tableStyle.Bold;
                }

                // Import the live formula, or a raw number, so the engine can compute/recompute; the
                // number format renders it (e.g. $50,000). Text/dates keep their formatted string.
                // Prefer the explicit format code; for built-in formats (percent, thousands, …) ClosedXML
                // leaves Format empty and only sets NumberFormatId, so map that to its implied code.
                string? format = importedStyle.NumberFormat;
                string text;
                string? numberFormat = null;
                string? formulaResult = null;
                double? formulaNumber = null;
                if (cell.HasFormula)
                {
                    text = "=" + StructuredReferences.Translate(cell.FormulaA1, r, c, tables);
                    numberFormat = format;
                    formulaResult = CachedFormulaText(cell.CachedValue, format);
                    formulaNumber = cell.CachedValue.Type == XLDataType.Number
                        ? cell.CachedValue.GetNumber() : null;
                }
                else if (cell.DataType == XLDataType.Number)
                {
                    text = cell.GetDouble().ToString(CultureInfo.InvariantCulture);
                    numberFormat = format;
                }
                else if (cell.DataType == XLDataType.Text)
                {
                    text = cell.GetString();
                }
                else if (cell.DataType == XLDataType.Boolean)
                {
                    text = cell.GetBoolean() ? "TRUE" : "FALSE";
                }
                else
                {
                    text = cell.GetFormattedString(); // dates/errors: show as Excel formatted them
                }

                row.Add(new CellData(
                    text,
                    bold,
                    importedStyle.Italic,
                    importedStyle.Underline,
                    foreground,
                    background,
                    importedStyle.FontSize,
                    AlignmentOf(importedStyle.HorizontalAlignment, cell.DataType),
                    numberFormat,
                    importedStyle.FontFamily,
                    importedStyle.Strikethrough,
                    formulaResult,
                    cell.HasComment ? cell.GetComment().Text : null,
                    cell.HasHyperlink
                        ? cell.GetHyperlink().ExternalAddress?.ToString() ??
                          (string.IsNullOrEmpty(cell.GetHyperlink().InternalAddress) ? null : "#" + cell.GetHyperlink().InternalAddress)
                        : null,
                    cell.HasComment ? cell.GetComment().Author : null,
                    formulaNumber));
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
        cancellationToken.ThrowIfCancellationRequested();
        List<SheetChart> charts = XlsxCharts.Read(bytes, sheet, sheetIndex);
        Dictionary<string, string> names = ReadNames(workbook, sheet);
        List<SheetDataBar> dataBars;
        List<SheetCellValueRule> cellValueRules;

        // Color-scale conditional formatting is not stored in each cell's base style. Paint its
        // current result into the imported cell snapshot so it appears in the editor and survives
        // a save even though the editor does not yet expose color-scale rule editing.
        using (ZipArchive archive = new(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read))
        {
            string partPath = WorksheetPartPath(archive, sheetIndex);
            IReadOnlyList<Color?> theme = ReadThemeColors(archive);
            (dataBars, cellValueRules) = ReadConditionalFormatting(archive, partPath, theme, cancellationToken);
            ApplyColorScales(rows, ReadColorScales(archive, partPath, theme), cancellationToken);
        }

        return new SheetData
        {
            Cells = rows,
            ColumnWidths = widths,
            Merges = merges,
            Images = images,
            Charts = charts,
            Names = names,
            DataBars = dataBars,
            CellValueRules = cellValueRules,
            AutoFilterRange = autoFilterRange,
            HiddenRows = hiddenRows,
            TabColor = sheet!.TabColor is { } tabColor && tabColor != XLColor.NoColor
                ? ResolveColor(workbook, tabColor)
                : null,
        };
    }

    // Read both layers of Excel's conditional formatting. In Hosts Inventory the data bar and the
    // red >90% font rule share K3:K602; both must survive preview, full load, edits and save.
    private static (List<SheetDataBar> Bars, List<SheetCellValueRule> ValueRules) ReadConditionalFormatting(
        ZipArchive archive, string partPath, IReadOnlyList<Color?> theme, CancellationToken cancellationToken)
    {
        List<SheetDataBar> bars = new();
        List<SheetCellValueRule> valueRules = new();
        ZipArchiveEntry? entry = archive.GetEntry(partPath);
        if (entry is null)
        {
            return (bars, valueRules);
        }

        List<(Color? Foreground, bool? Bold)>? fonts = null;
        using Stream stream = entry.Open();
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = true,
        });
        while (!reader.EOF)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }
            if (reader.LocalName == "sheetData")
            {
                reader.Skip();
                continue;
            }
            if (reader.LocalName != "conditionalFormatting")
            {
                reader.Read();
                continue;
            }

            XElement block = (XElement)XNode.ReadFrom(reader);
            List<CellRange> ranges = (block.Attribute("sqref")?.Value ?? string.Empty)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries)
                .Select(static reference => CellRange.TryParse(reference, out CellRange range) ? (CellRange?)range : null)
                .Where(static range => range is not null).Select(static range => range!.Value).ToList();
            foreach (XElement rule in block.Elements(Ssml + "cfRule"))
            {
                XElement? dataBar = rule.Element(Ssml + "dataBar");
                if (dataBar is not null)
                {
                    XElement[] limits = dataBar.Elements(Ssml + "cfvo").ToArray();
                    Color? color = PreviewColor(dataBar.Element(Ssml + "color"), theme);
                    if (limits.Length < 2 || color is null)
                    {
                        continue;
                    }
                    double min = double.TryParse(limits[0].Attribute("val")?.Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out double parsedMin) ? parsedMin : 0;
                    double max = double.TryParse(limits[1].Attribute("val")?.Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out double parsedMax) ? parsedMax : 1;
                    foreach (CellRange range in ranges)
                    {
                        bars.Add(new SheetDataBar
                        {
                            Range = range,
                            Min = min,
                            Max = max,
                            Color = color.Value,
                            ShowValue = dataBar.Attribute("showValue")?.Value is not ("0" or "false"),
                            Gradient = dataBar.Attribute("gradient")?.Value is not ("0" or "false"),
                        });
                    }
                    continue;
                }

                if (rule.Attribute("type")?.Value != "cellIs" ||
                    !TryComparison(rule.Attribute("operator")?.Value, out CellValueComparison comparison) ||
                    !double.TryParse(rule.Element(Ssml + "formula")?.Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out double threshold) ||
                    !int.TryParse(rule.Attribute("dxfId")?.Value, out int dxfId) || dxfId < 0)
                {
                    continue;
                }
                fonts ??= ReadDifferentialFonts(archive, theme);
                if (dxfId >= fonts.Count)
                {
                    continue;
                }
                (Color? foreground, bool? bold) = fonts[dxfId];
                if (foreground is null && bold is null)
                {
                    continue;
                }
                int priority = int.TryParse(rule.Attribute("priority")?.Value, out int parsedPriority)
                    ? parsedPriority : int.MaxValue;
                foreach (CellRange range in ranges)
                {
                    valueRules.Add(new SheetCellValueRule
                    {
                        Range = range,
                        Comparison = comparison,
                        Threshold = threshold,
                        Foreground = foreground,
                        Bold = bold,
                        Priority = priority,
                    });
                }
            }
        }
        valueRules.Sort(static (left, right) => left.Priority.CompareTo(right.Priority));
        return (bars, valueRules);
    }

    private static List<(Color? Foreground, bool? Bold)> ReadDifferentialFonts(
        ZipArchive archive, IReadOnlyList<Color?> theme)
    {
        List<(Color?, bool?)> styles = new();
        if (archive.GetEntry("xl/styles.xml") is not { } entry)
        {
            return styles;
        }
        using Stream stream = entry.Open();
        XDocument document = XDocument.Load(stream);
        foreach (XElement dxf in document.Descendants(Ssml + "dxfs").Elements(Ssml + "dxf"))
        {
            XElement? font = dxf.Element(Ssml + "font");
            XElement? bold = font?.Element(Ssml + "b");
            styles.Add((PreviewColor(font?.Element(Ssml + "color"), theme),
                bold is null ? null : bold.Attribute("val")?.Value is not ("0" or "false")));
        }
        return styles;
    }

    private static bool TryComparison(string? value, out CellValueComparison comparison)
    {
        comparison = value switch
        {
            "greaterThan" => CellValueComparison.GreaterThan,
            "greaterThanOrEqual" => CellValueComparison.GreaterThanOrEqual,
            "lessThan" => CellValueComparison.LessThan,
            "lessThanOrEqual" => CellValueComparison.LessThanOrEqual,
            "equal" => CellValueComparison.Equal,
            "notEqual" => CellValueComparison.NotEqual,
            _ => default,
        };
        return value is "greaterThan" or "greaterThanOrEqual" or "lessThan" or
            "lessThanOrEqual" or "equal" or "notEqual";
    }

    private readonly record struct ColorScaleStop(string Type, double Value, Color Color);
    private readonly record struct ColorScaleRule(CellRange Range, IReadOnlyList<ColorScaleStop> Stops);

    private static string WorksheetPartPath(ZipArchive archive, int sheetIndex)
    {
        XDocument workbook = ReadXmlPart(archive, "xl/workbook.xml");
        XDocument relationships = ReadXmlPart(archive, "xl/_rels/workbook.xml.rels");
        XNamespace officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement sheet = workbook.Descendants(Ssml + "sheet").ElementAt(sheetIndex);
        string id = sheet.Attribute(officeRelationships + "id")?.Value ?? string.Empty;
        XElement relationship = relationships.Descendants(packageRelationships + "Relationship")
            .First(r => r.Attribute("Id")?.Value == id);
        return ResolvePartPath("xl/workbook.xml", relationship.Attribute("Target")!.Value);
    }

    // Unlike ordinary fills, Excel color scales live in conditionalFormatting after sheetData.
    // Stream past the cells to avoid allocating a second full worksheet DOM.
    private static List<ColorScaleRule> ReadColorScales(
        ZipArchive archive, string partPath, IReadOnlyList<Color?> theme)
    {
        List<ColorScaleRule> rules = new();
        ZipArchiveEntry? entry = archive.GetEntry(partPath);
        if (entry is null)
        {
            return rules;
        }

        using Stream stream = entry.Open();
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = true,
        });
        while (!reader.EOF)
        {
            if (reader.NodeType != XmlNodeType.Element)
            {
                reader.Read();
                continue;
            }

            if (reader.LocalName == "sheetData")
            {
                reader.Skip();
                continue;
            }

            if (reader.LocalName != "conditionalFormatting")
            {
                reader.Read();
                continue;
            }

            XElement block = (XElement)XNode.ReadFrom(reader);
            string[] references = (block.Attribute("sqref")?.Value ?? string.Empty)
                .Split(' ', StringSplitOptions.RemoveEmptyEntries);
            foreach (XElement scale in block.Elements(Ssml + "cfRule").Elements(Ssml + "colorScale"))
            {
                XElement[] thresholds = scale.Elements(Ssml + "cfvo").ToArray();
                XElement[] colors = scale.Elements(Ssml + "color").ToArray();
                if (thresholds.Length is < 2 or > 3 || colors.Length != thresholds.Length)
                {
                    continue;
                }

                List<ColorScaleStop> stops = new();
                for (int i = 0; i < thresholds.Length; i++)
                {
                    Color? color = PreviewColor(colors[i], theme);
                    if (color is null)
                    {
                        break;
                    }
                    double.TryParse(thresholds[i].Attribute("val")?.Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out double value);
                    stops.Add(new ColorScaleStop(thresholds[i].Attribute("type")?.Value ?? "num", value, color.Value));
                }

                if (stops.Count != thresholds.Length)
                {
                    continue;
                }
                foreach (string reference in references)
                {
                    if (CellRange.TryParse(reference, out CellRange range))
                    {
                        rules.Add(new ColorScaleRule(range, stops));
                    }
                }
            }
        }
        return rules;
    }

    private static void ApplyColorScales(
        List<List<CellData>> rows, IReadOnlyList<ColorScaleRule> rules, CancellationToken cancellationToken)
    {
        foreach (ColorScaleRule rule in rules)
        {
            List<double> numbers = new();
            int lastRow = Math.Min(rows.Count, rule.Range.Row + rule.Range.RowSpan);
            for (int row = Math.Max(0, rule.Range.Row); row < lastRow; row++)
            {
                for (int column = Math.Max(0, rule.Range.Column);
                     column < Math.Min(rows[row].Count, rule.Range.Column + rule.Range.ColumnSpan); column++)
                {
                    CellData cell = rows[row][column];
                    if (double.TryParse(cell.FormulaResult ?? cell.Text, NumberStyles.Any,
                        CultureInfo.InvariantCulture, out double number))
                    {
                        numbers.Add(number);
                    }
                }
            }
            if (numbers.Count == 0)
            {
                continue;
            }

            numbers.Sort();
            double min = numbers[0], max = numbers[^1];
            double Threshold(ColorScaleStop stop) => stop.Type.ToLowerInvariant() switch
            {
                "min" => min,
                "max" => max,
                "percent" => min + ((max - min) * stop.Value / 100),
                "percentile" => Percentile(numbers, stop.Value),
                _ => stop.Value,
            };
            double[] thresholds = rule.Stops.Select(Threshold).ToArray();
            for (int row = Math.Max(0, rule.Range.Row); row < lastRow; row++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                for (int column = Math.Max(0, rule.Range.Column);
                     column < Math.Min(rows[row].Count, rule.Range.Column + rule.Range.ColumnSpan); column++)
                {
                    CellData cell = rows[row][column];
                    if (!double.TryParse(cell.FormulaResult ?? cell.Text, NumberStyles.Any,
                        CultureInfo.InvariantCulture, out double number))
                    {
                        continue;
                    }
                    int upper = number <= thresholds[1] ? 1 : rule.Stops.Count - 1;
                    double span = thresholds[upper] - thresholds[upper - 1];
                    double fraction = span == 0 ? 1 : Math.Clamp((number - thresholds[upper - 1]) / span, 0, 1);
                    rows[row][column] = cell with { Background = Interpolate(rule.Stops[upper - 1].Color, rule.Stops[upper].Color, fraction) };
                }
            }
        }
    }

    private static double Percentile(IReadOnlyList<double> sorted, double percent)
    {
        double position = Math.Clamp(percent / 100, 0, 1) * (sorted.Count - 1);
        int lower = (int)position;
        int upper = Math.Min(sorted.Count - 1, lower + 1);
        return sorted[lower] + ((sorted[upper] - sorted[lower]) * (position - lower));
    }

    private static Color Interpolate(Color first, Color second, double amount)
    {
        byte Mix(byte a, byte b) => (byte)Math.Round(a + ((b - a) * amount));
        return Color.FromArgb(Mix(first.A, second.A), Mix(first.R, second.R),
            Mix(first.G, second.G), Mix(first.B, second.B));
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

    // Applies an Excel tint (-1..1) in the HLS color space Excel uses. A per-channel linear blend
    // shifts saturated theme colors and produces noticeably wrong table stripes.
    private static Color ApplyTint(Color c, double tint)
    {
        if (tint == 0)
        {
            return c;
        }

        const int hlsMax = 240;
        const int rgbMax = 255;

        int red = c.R;
        int green = c.G;
        int blue = c.B;
        int maximum = Math.Max(red, Math.Max(green, blue));
        int minimum = Math.Min(red, Math.Min(green, blue));
        int sum = maximum + minimum;
        int delta = maximum - minimum;
        int luminance = (sum * hlsMax + rgbMax) / (2 * rgbMax);
        int hue;
        int saturation;

        if (delta == 0)
        {
            hue = hlsMax * 2 / 3;
            saturation = 0;
        }
        else
        {
            saturation = luminance <= hlsMax / 2
                ? (delta * hlsMax + sum / 2) / sum
                : (delta * hlsMax + (2 * rgbMax - sum) / 2) / (2 * rgbMax - sum);

            int redDelta = ((maximum - red) * (hlsMax / 6) + delta / 2) / delta;
            int greenDelta = ((maximum - green) * (hlsMax / 6) + delta / 2) / delta;
            int blueDelta = ((maximum - blue) * (hlsMax / 6) + delta / 2) / delta;
            hue = maximum == red
                ? blueDelta - greenDelta
                : maximum == green
                    ? hlsMax / 3 + redDelta - blueDelta
                    : 2 * hlsMax / 3 + greenDelta - redDelta;
            if (hue < 0)
            {
                hue += hlsMax;
            }
            else if (hue > hlsMax)
            {
                hue -= hlsMax;
            }
        }

        luminance = (int)Math.Round(
            tint < 0 ? luminance * (1.0 + tint) : luminance * (1.0 - tint) + hlsMax * tint,
            MidpointRounding.AwayFromZero);
        luminance = Math.Clamp(luminance, 0, hlsMax);

        if (saturation == 0)
        {
            byte gray = (byte)((luminance * rgbMax + hlsMax / 2) / hlsMax);
            return Color.FromArgb(c.A, gray, gray, gray);
        }

        int magic2 = luminance <= hlsMax / 2
            ? (luminance * (hlsMax + saturation) + hlsMax / 2) / hlsMax
            : luminance + saturation - (luminance * saturation + hlsMax / 2) / hlsMax;
        int magic1 = 2 * luminance - magic2;

        byte ToRgb(int hueValue)
        {
            if (hueValue < 0)
            {
                hueValue += hlsMax;
            }
            else if (hueValue > hlsMax)
            {
                hueValue -= hlsMax;
            }

            int value;
            if (hueValue < hlsMax / 6)
            {
                value = magic1 + ((magic2 - magic1) * hueValue + hlsMax / 12) / (hlsMax / 6);
            }
            else if (hueValue < hlsMax / 2)
            {
                value = magic2;
            }
            else if (hueValue < 2 * hlsMax / 3)
            {
                value = magic1 + ((magic2 - magic1) * (2 * hlsMax / 3 - hueValue) + hlsMax / 12) / (hlsMax / 6);
            }
            else
            {
                value = magic1;
            }

            return (byte)Math.Clamp((value * rgbMax + hlsMax / 2) / hlsMax, 0, rgbMax);
        }

        return Color.FromArgb(c.A, ToRgb(hue + hlsMax / 3), ToRgb(hue), ToRgb(hue - hlsMax / 3));
    }

    // ----- Excel Table formatting (header / total / striped rows from the table style) -----

    private static readonly XNamespace Ssml = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    // Builds a per-cell style map for the sheet's tables. Excel stores custom table styles in
    // styles.xml, but its built-in styles are referenced only by name and must be resolved from the
    // workbook theme.
    private static Dictionary<(int Row, int Column), TableCellStyle> BuildTableStyles(byte[] bytes, IXLWorksheet? sheet, XLWorkbook workbook)
    {
        Dictionary<(int, int), TableCellStyle> cells = new();
        if (sheet is null || !sheet.Tables.Any())
        {
            return cells;
        }

        try
        {
            // Resolve the workbook's custom table-style definitions (style name → element DXFs).
            using ZipArchive archive = new(new MemoryStream(bytes), ZipArchiveMode.Read);
            ZipArchiveEntry? stylesEntry = archive.GetEntry("xl/styles.xml");
            if (stylesEntry is null)
            {
                return cells;
            }

            XDocument styles = XDocument.Load(stylesEntry.Open());
            List<XElement> dxfs = styles.Descendants(Ssml + "dxfs").Elements(Ssml + "dxf").ToList();

            Dictionary<string, ResolvedTableStyle> tableStyles = new(StringComparer.OrdinalIgnoreCase);
            foreach (XElement ts in styles.Descendants(Ssml + "tableStyle"))
            {
                string? name = ts.Attribute("name")?.Value;
                if (string.IsNullOrEmpty(name))
                {
                    continue;
                }

                TableCellStyle Element(string type, out int size)
                {
                    size = 1;
                    XElement? el = ts.Elements(Ssml + "tableStyleElement").FirstOrDefault(e => e.Attribute("type")?.Value == type);
                    if (el is null || !int.TryParse(el.Attribute("dxfId")?.Value, out int id) || id < 0 || id >= dxfs.Count)
                    {
                        return default;
                    }

                    if (int.TryParse(el.Attribute("size")?.Value, out int s) && s > 0)
                    {
                        size = s;
                    }

                    return DxfStyle(dxfs[id], workbook);
                }

                TableCellStyle whole = Element("wholeTable", out _);
                TableCellStyle header = Element("headerRow", out _);
                TableCellStyle total = Element("totalRow", out _);
                TableCellStyle stripe1 = Element("firstRowStripe", out int size1);
                TableCellStyle stripe2 = Element("secondRowStripe", out int size2);
                tableStyles[name] = new ResolvedTableStyle(whole, header, total, stripe1, stripe2, size1, size2);
            }

            foreach (IXLTable table in sheet.Tables)
            {
                string? styleName = table.Theme?.Name;
                if (styleName is null ||
                    (!tableStyles.TryGetValue(styleName, out ResolvedTableStyle style) &&
                     !TryBuiltInTableStyle(workbook, styleName, out style)))
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

                void Paint(int row, TableCellStyle cellStyle)
                {
                    if (row < 1 || cellStyle == default)
                    {
                        return;
                    }

                    for (int col = firstColumn; col <= lastColumn; col++)
                    {
                        (int Row, int Column) key = (row, col);
                        cells.TryGetValue(key, out TableCellStyle current);
                        cells[key] = new TableCellStyle(
                            cellStyle.Background ?? current.Background,
                            cellStyle.Foreground ?? current.Foreground,
                            current.Bold || cellStyle.Bold);
                    }
                }

                for (int row = range.FirstAddress.RowNumber; row <= range.LastAddress.RowNumber; row++)
                {
                    Paint(row, style.Whole);
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
            // Table formatting is cosmetic — never let it break opening the file.
        }

        return cells;
    }

    private readonly record struct TableCellStyle(Color? Background, Color? Foreground, bool Bold);

    private readonly record struct ImportedCellStyle(
        bool Bold,
        bool Italic,
        bool Underline,
        bool Strikethrough,
        Color? Foreground,
        Color? Background,
        double FontSize,
        string? FontFamily,
        string? NumberFormat,
        XLAlignmentHorizontalValues HorizontalAlignment);

    private static string CachedFormulaText(XLCellValue value, string? numberFormat) => value.Type switch
    {
        XLDataType.Blank => string.Empty,
        XLDataType.Boolean => value.GetBoolean() ? "TRUE" : "FALSE",
        XLDataType.Number => NumberFormatter.Format(value.GetNumber(), numberFormat),
        XLDataType.Text => value.GetText(),
        XLDataType.Error => value.GetError() switch
        {
            XLError.NullValue => "#NULL!",
            XLError.DivisionByZero => "#DIV/0!",
            XLError.IncompatibleValue => "#VALUE!",
            XLError.CellReference => "#REF!",
            XLError.NameNotRecognized => "#NAME?",
            XLError.NumberInvalid => "#NUM!",
            XLError.NoValueAvailable => "#N/A",
            _ => "#VALUE!",
        },
        XLDataType.DateTime => value.GetDateTime().ToString(CultureInfo.InvariantCulture),
        XLDataType.TimeSpan => value.GetTimeSpan().ToString(),
        _ => value.ToString(CultureInfo.InvariantCulture),
    };

    private static ImportedCellStyle ReadCellStyle(XLWorkbook workbook, IXLStyle style)
    {
        IXLFontBase font = style.Font;

        // Drop default black so plain cells stay unstyled and a table style can supply its foreground.
        Color? foreground = ResolveColor(workbook, font.FontColor);
        if (foreground is { } color && color.R == 0 && color.G == 0 && color.B == 0)
        {
            foreground = null;
        }

        Color? background = style.Fill.PatternType != XLFillPatternValues.None
            ? ResolveColor(workbook, style.Fill.BackgroundColor)
            : null;
        string? numberFormat = style.NumberFormat.Format;
        numberFormat = string.IsNullOrWhiteSpace(numberFormat)
            ? BuiltInNumberFormat(style.NumberFormat.NumberFormatId)
            : numberFormat;

        return new ImportedCellStyle(
            font.Bold,
            font.Italic,
            font.Underline != XLFontUnderlineValues.None,
            font.Strikethrough,
            foreground,
            background,
            font.FontSize * 4.0 / 3.0, // points -> device pixels (96/72)
            string.IsNullOrWhiteSpace(font.FontName) ? null : font.FontName,
            numberFormat,
            style.Alignment.Horizontal);
    }

    private readonly record struct ResolvedTableStyle(
        TableCellStyle Whole,
        TableCellStyle Header,
        TableCellStyle Total,
        TableCellStyle Stripe1,
        TableCellStyle Stripe2,
        int Stripe1Size,
        int Stripe2Size);

    // Excel's 60 built-in table styles are not serialized as DXFs. Their colors follow the workbook
    // theme in groups of seven (Text1, Accent1..Accent6), with four medium and three light variants.
    private static bool TryBuiltInTableStyle(XLWorkbook workbook, string name, out ResolvedTableStyle style)
    {
        style = default;
        string family;
        int maximum;
        if (name.StartsWith("TableStyleLight", StringComparison.OrdinalIgnoreCase))
        {
            family = "Light";
            maximum = 21;
        }
        else if (name.StartsWith("TableStyleMedium", StringComparison.OrdinalIgnoreCase))
        {
            family = "Medium";
            maximum = 28;
        }
        else if (name.StartsWith("TableStyleDark", StringComparison.OrdinalIgnoreCase))
        {
            family = "Dark";
            maximum = 11;
        }
        else
        {
            return false;
        }

        string numberText = name[("TableStyle".Length + family.Length)..];
        if (!int.TryParse(numberText, NumberStyles.None, CultureInfo.InvariantCulture, out int number) ||
            number < 1 || number > maximum)
        {
            return false;
        }

        Color black = ThemeTableColor(workbook, 0);
        Color white = Color.FromRgb(0xFF, 0xFF, 0xFF);
        int slot = (number - 1) % 7;
        Color baseColor = ThemeTableColor(workbook, slot);
        Color light = TableTint(baseColor, slot, slot == 0 ? 0.85 : 0.80);
        Color mediumLight = TableTint(baseColor, slot, slot == 0 ? 0.65 : 0.60);
        Color dark = slot == 0 ? black : TableTint(baseColor, slot, -0.25);

        TableCellStyle None(Color? foreground = null, bool bold = false) => new(null, foreground, bold);
        TableCellStyle Fill(Color background, Color? foreground = null, bool bold = false) => new(background, foreground, bold);

        int group = (number - 1) / 7;
        switch (family)
        {
            case "Light" when group == 0:
                style = new ResolvedTableStyle(None(dark), None(dark, true), None(dark, true), Fill(light, dark), None(dark), 1, 1);
                return true;
            case "Light" when group == 1:
                style = new ResolvedTableStyle(default, Fill(baseColor, white, true), None(black, true), default, default, 1, 1);
                return true;
            case "Light":
                style = new ResolvedTableStyle(default, None(black, true), None(black, true), Fill(light, black), default, 1, 1);
                return true;
            case "Medium" when group == 0:
                style = new ResolvedTableStyle(default, Fill(baseColor, white, true), None(black, true), Fill(light, black), default, 1, 1);
                return true;
            case "Medium" when group == 1:
                style = new ResolvedTableStyle(default, Fill(baseColor, white, true), Fill(baseColor, white, true), Fill(mediumLight, black), Fill(light, black), 1, 1);
                return true;
            case "Medium" when group == 2:
                style = new ResolvedTableStyle(default, Fill(baseColor, white, true), None(black), Fill(TableTint(black, 0, 0.85), black), default, 1, 1);
                return true;
            case "Medium":
                style = new ResolvedTableStyle(default, Fill(light, black, true), Fill(light, black, true), Fill(mediumLight, black), Fill(light, black), 1, 1);
                return true;
            case "Dark" when number <= 7:
                {
                    Color stripe1 = TableTint(baseColor, slot, slot == 0 ? 0.25 : -0.25);
                    Color stripe2 = slot == 0 ? TableTint(baseColor, slot, 0.45) : baseColor;
                    Color total = TableTint(baseColor, slot, slot == 0 ? 0.15 : -0.50);
                    style = new ResolvedTableStyle(None(white), Fill(black, white, true), Fill(total, white, true), Fill(stripe1, white), Fill(stripe2, white), 1, 1);
                    return true;
                }
            case "Dark" when number == 8:
                style = new ResolvedTableStyle(default, Fill(black, white), Fill(TableTint(black, 0, 0.85), black, true), Fill(TableTint(black, 0, 0.65), black), Fill(TableTint(black, 0, 0.85), black), 1, 1);
                return true;
            case "Dark":
                {
                    int accentPair = (number - 9) * 2 + 1;
                    Color stripeBase = ThemeTableColor(workbook, accentPair);
                    Color header = ThemeTableColor(workbook, accentPair + 1);
                    Color stripeLight = TableTint(stripeBase, accentPair, 0.80);
                    style = new ResolvedTableStyle(default, Fill(header, white), Fill(stripeLight, black, true), Fill(TableTint(stripeBase, accentPair, 0.60), black), Fill(stripeLight, black), 1, 1);
                    return true;
                }
        }

        return false;
    }

    private static Color ThemeTableColor(XLWorkbook workbook, int slot)
    {
        XLThemeColor theme = slot switch
        {
            1 => XLThemeColor.Accent1,
            2 => XLThemeColor.Accent2,
            3 => XLThemeColor.Accent3,
            4 => XLThemeColor.Accent4,
            5 => XLThemeColor.Accent5,
            6 => XLThemeColor.Accent6,
            _ => XLThemeColor.Text1,
        };
        Color fallback = DefaultTableColor(slot);
        XLColor? themeColor = ThemeBase(workbook, theme);
        return themeColor is null ? fallback : ResolveColor(workbook, themeColor) ?? fallback;
    }

    private static Color DefaultTableColor(int slot) => slot switch
    {
        1 => Color.FromRgb(0x44, 0x72, 0xC4),
        2 => Color.FromRgb(0xED, 0x7D, 0x31),
        3 => Color.FromRgb(0xA5, 0xA5, 0xA5),
        4 => Color.FromRgb(0xFF, 0xC0, 0x00),
        5 => Color.FromRgb(0x5B, 0x9B, 0xD5),
        6 => Color.FromRgb(0x70, 0xAD, 0x47),
        _ => Color.FromRgb(0x00, 0x00, 0x00),
    };

    // Excel's standard Office theme uses a fixed, well-known table-style palette. Preserve those
    // exact RGB values; for a custom workbook theme, calculate the same relative tint instead.
    private static Color TableTint(Color color, int slot, double tint)
    {
        if (color == DefaultTableColor(slot))
        {
            string? hex = (slot, tint) switch
            {
                (0, 0.85) => "D9D9D9",
                (0, 0.65) => "A6A6A6",
                (0, 0.25) => "404040",
                (0, 0.45) => "737373",
                (0, 0.15) => "262626",
                (1, 0.80) => "D9E1F2",
                (1, 0.60) => "B4C6E7",
                (1, -0.25) => "305496",
                (1, -0.50) => "203764",
                (2, 0.80) => "FCE4D6",
                (2, 0.60) => "F8CBAD",
                (2, -0.25) => "C65911",
                (2, -0.50) => "833C0C",
                (3, 0.80) => "EDEDED",
                (3, 0.60) => "DBDBDB",
                (3, -0.25) => "7B7B7B",
                (3, -0.50) => "525252",
                (4, 0.80) => "FFF2CC",
                (4, 0.60) => "FFE699",
                (4, -0.25) => "BF8F00",
                (4, -0.50) => "806000",
                (5, 0.80) => "DDEBF7",
                (5, 0.60) => "BDD7EE",
                (5, -0.25) => "2F75B5",
                (5, -0.50) => "1F4E78",
                (6, 0.80) => "E2EFDA",
                (6, 0.60) => "C6E0B4",
                (6, -0.25) => "548235",
                (6, -0.50) => "375623",
                _ => null,
            };
            if (hex is not null)
            {
                return Color.FromRgb(
                    Convert.ToByte(hex[..2], 16),
                    Convert.ToByte(hex.Substring(2, 2), 16),
                    Convert.ToByte(hex.Substring(4, 2), 16));
            }
        }

        return ApplyTint(color, tint);
    }

    private static TableCellStyle DxfStyle(XElement dxf, XLWorkbook workbook)
    {
        XElement? font = dxf.Element(Ssml + "font");
        return new TableCellStyle(
            DxfFill(dxf, workbook),
            DxfColor(font?.Element(Ssml + "color"), workbook),
            font?.Element(Ssml + "b") is not null);
    }

    // Resolves a dxf fill. Solid SpreadsheetML fills use fgColor; bgColor is for the secondary
    // color of patterned fills, so it is only a fallback here.
    private static Color? DxfFill(XElement dxf, XLWorkbook workbook)
    {
        XElement? pattern = dxf.Element(Ssml + "fill")?.Element(Ssml + "patternFill");
        XElement? color = pattern?.Element(Ssml + "fgColor") ?? pattern?.Element(Ssml + "bgColor");
        return DxfColor(color, workbook);
    }

    private static Color? DxfColor(XElement? element, XLWorkbook workbook)
    {
        if (element is null)
        {
            return null;
        }

        string? rgb = element.Attribute("rgb")?.Value;
        if (!string.IsNullOrEmpty(rgb) && rgb.Length >= 6)
        {
            string hex = rgb.Length == 8 ? rgb : "FF" + rgb;
            return Color.FromArgb(
                Convert.ToByte(hex[..2], 16), Convert.ToByte(hex.Substring(2, 2), 16),
                Convert.ToByte(hex.Substring(4, 2), 16), Convert.ToByte(hex.Substring(6, 2), 16));
        }

        if (int.TryParse(element.Attribute("theme")?.Value, out int themeIndex))
        {
            double tint = double.TryParse(element.Attribute("tint")?.Value, NumberStyles.Any, CultureInfo.InvariantCulture, out double t) ? t : 0;
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

    // Compatibility helper for hosts that save one worksheet directly.
    public static void Write(
        Stream stream,
        IReadOnlyList<IReadOnlyList<CellData>> rows,
        IReadOnlyDictionary<string, string>? names = null,
        IReadOnlyList<MergeRange>? merges = null,
        string? sheetName = null,
        IReadOnlyList<SheetDataBar>? dataBars = null,
        IReadOnlyList<ChartWriteData>? charts = null,
        IReadOnlyList<double>? columnWidths = null,
        Color? tabColor = null)
    {
        SheetData data = new()
        {
            Cells = rows,
            Names = names ?? new Dictionary<string, string>(),
            Merges = merges ?? Array.Empty<MergeRange>(),
            DataBars = dataBars ?? Array.Empty<SheetDataBar>(),
            ColumnWidths = columnWidths ?? Array.Empty<double>(),
            TabColor = tabColor,
        };
        WriteWorkbook(
            stream,
            new[] { new XlsxSheetWriteData(sheetName ?? EditorText.Get(EditorString.Sheet) + "1", data, charts ?? Array.Empty<ChartWriteData>()) });
    }

    // Writes the complete workbook in tab order. Each worksheet keeps its cells, styles, dimensions,
    // merges, names, data bars, tab color, and chart definitions.
    internal static void WriteWorkbook(Stream stream, IReadOnlyList<XlsxSheetWriteData> sheets)
    {
        using XLWorkbook workbook = new();
        IReadOnlyList<XlsxSheetWriteData> source = sheets.Count > 0
            ? sheets
            : new[]
            {
                new XlsxSheetWriteData(
                    EditorText.Get(EditorString.Sheet) + "1",
                    new SheetData { Cells = new[] { new[] { CellData.Plain(string.Empty) } } },
                    Array.Empty<ChartWriteData>()),
            };

        List<(int SheetIndex, IReadOnlyList<ChartWriteData> Charts)> chartGroups = new();
        for (int index = 0; index < source.Count; index++)
        {
            XlsxSheetWriteData input = source[index];
            IXLWorksheet sheet = workbook.Worksheets.Add(UniqueSheetName(workbook, input.Name));
            if (input.Data.TabColor is { } tabColor)
            {
                sheet.TabColor = ToXl(tabColor);
            }

            WriteSheet(workbook, sheet, input.Data);
            if (input.Charts.Count > 0)
            {
                chartGroups.Add((index, input.Charts));
            }
        }

        if (chartGroups.Count == 0)
        {
            workbook.SaveAs(stream);
            return;
        }

        using MemoryStream buffer = new();
        workbook.SaveAs(buffer);
        byte[] bytes = buffer.ToArray();
        foreach ((int sheetIndex, IReadOnlyList<ChartWriteData> charts) in chartGroups)
        {
            bytes = XlsxChartWriter.Inject(bytes, charts, sheetIndex);
        }

        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteSheet(XLWorkbook workbook, IXLWorksheet sheet, SheetData data)
    {
        IReadOnlyList<IReadOnlyList<CellData>> rows = data.Cells;
        IReadOnlyList<double> columnWidths = data.ColumnWidths;
        IReadOnlyList<MergeRange> merges = data.Merges;
        IReadOnlyList<SheetDataBar> dataBars = data.DataBars;
        IReadOnlyList<SheetCellValueRule> valueRules = data.CellValueRules;
        CellRange? autoFilterRange = data.AutoFilterRange;
        IReadOnlyList<int> hiddenRows = data.HiddenRows;
        IReadOnlyDictionary<string, string> names = data.Names;

        for (int r = 0; r < rows.Count; r++)
        {
            for (int c = 0; c < rows[r].Count; c++)
            {
                CellData cellData = rows[r][c];
                IXLCell cell = sheet.Cell(r + 1, c + 1);

                if (cellData.Text.StartsWith('=') && cellData.Text.Length > 1)
                {
                    cell.FormulaA1 = cellData.Text[1..];          // round-trip formulas
                }
                else if (double.TryParse(cellData.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double num))
                {
                    cell.Value = num;                          // numbers as numbers (so they stay computable)
                }
                else if (!string.IsNullOrEmpty(cellData.Text))
                {
                    cell.Value = cellData.Text;
                }

                if (!string.IsNullOrWhiteSpace(cellData.Comment))
                {
                    IXLComment comment = cell.CreateComment();
                    comment.AddText(cellData.Comment);
                    comment.Author = string.IsNullOrWhiteSpace(cellData.CommentAuthor)
                        ? "Spreadsheet Editor"
                        : cellData.CommentAuthor;
                }

                if (!string.IsNullOrWhiteSpace(cellData.Hyperlink))
                {
                    string target = cellData.Hyperlink;
                    cell.SetHyperlink(target.StartsWith('#')
                        ? new XLHyperlink(target[1..])
                        : new XLHyperlink(new Uri(target, UriKind.Absolute)));
                }

                if (cellData.NumberFormat is { } nf)
                {
                    cell.Style.NumberFormat.Format = nf;
                }

                if (cellData.FontFamily is { } fam)
                {
                    cell.Style.Font.FontName = fam;
                }

                if (cellData.Bold)
                {
                    cell.Style.Font.Bold = true;
                }

                if (cellData.Italic)
                {
                    cell.Style.Font.Italic = true;
                }

                if (cellData.Underline)
                {
                    cell.Style.Font.Underline = XLFontUnderlineValues.Single;
                }

                if (cellData.Strikethrough)
                {
                    cell.Style.Font.Strikethrough = true;
                }

                if (cellData.Foreground is { } fg)
                {
                    cell.Style.Font.FontColor = ToXl(fg);
                }

                if (cellData.Background is { } bg)
                {
                    cell.Style.Fill.PatternType = XLFillPatternValues.Solid;
                    cell.Style.Fill.BackgroundColor = ToXl(bg);
                }

                if (cellData.FontSize > 0)
                {
                    cell.Style.Font.FontSize = cellData.FontSize * 3.0 / 4.0; // px -> points
                }

                cell.Style.Alignment.Horizontal = cellData.Alignment switch
                {
                    TextAlignment.Center => XLAlignmentHorizontalValues.Center,
                    TextAlignment.Right => XLAlignmentHorizontalValues.Right,
                    _ => XLAlignmentHorizontalValues.Left,
                };
            }
        }

        foreach (int row in hiddenRows)
        {
            if (row >= 0 && row < rows.Count)
            {
                sheet.Row(row + 1).Hide();
            }
        }

        if (autoFilterRange is { RowSpan: > 0, ColumnSpan: > 0 } filter)
        {
            try
            {
                sheet.Range(filter.Row + 1, filter.Column + 1,
                    filter.Row + filter.RowSpan, filter.Column + filter.ColumnSpan).SetAutoFilter();
            }
            catch
            {
                // Keep the worksheet writable even if a source range is invalid after editing.
            }
        }

        // Persist column widths (px) so the layout survives the round-trip — otherwise every column
        // falls back to Excel's default and the content truncates. Invert the reader's chars->px
        // conversion (px = chars * 7 + 5).
        if (columnWidths.Count > 0)
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
        if (merges.Count > 0)
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
        if (dataBars.Count > 0)
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

        foreach (SheetCellValueRule rule in valueRules.OrderBy(static rule => rule.Priority))
        {
            if (rule.Range.RowSpan <= 0 || rule.Range.ColumnSpan <= 0)
            {
                continue;
            }
            IXLRange range = sheet.Range(
                rule.Range.Row + 1, rule.Range.Column + 1,
                rule.Range.Row + rule.Range.RowSpan, rule.Range.Column + rule.Range.ColumnSpan);
            IXLConditionalFormat conditional = range.AddConditionalFormat();
            IXLStyle style = rule.Comparison switch
            {
                CellValueComparison.GreaterThan => conditional.WhenGreaterThan(rule.Threshold),
                CellValueComparison.GreaterThanOrEqual => conditional.WhenEqualOrGreaterThan(rule.Threshold),
                CellValueComparison.LessThan => conditional.WhenLessThan(rule.Threshold),
                CellValueComparison.LessThanOrEqual => conditional.WhenEqualOrLessThan(rule.Threshold),
                CellValueComparison.Equal => conditional.WhenEquals(rule.Threshold),
                CellValueComparison.NotEqual => conditional.WhenNotEquals(rule.Threshold),
                _ => throw new InvalidDataException(EditorText.Get(EditorString.UnsupportedConditionalComparison)),
            };
            if (rule.Foreground is { } foreground)
            {
                style.Font.FontColor = ToXl(foreground);
            }
            if (rule.Bold is { } bold)
            {
                style.Font.Bold = bold;
            }
        }

        // Persist workbook defined names so formulas that reference a name (=NetIncome, etc.) still
        // resolve on the next open — otherwise they come back as #NAME?. A bare range ("$B$1") isn't a
        // valid defined-name target, so qualify it with the sheet; fall back to the raw value.
        if (names.Count > 0)
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

    }

    private static string UniqueSheetName(XLWorkbook workbook, string? desired)
    {
        string baseName = SafeSheetName(desired);
        string candidate = baseName;
        for (int suffix = 2; workbook.Worksheets.Any(sheet => sheet.Name.Equals(candidate, StringComparison.OrdinalIgnoreCase)); suffix++)
        {
            string marker = $" ({suffix})";
            candidate = baseName[..Math.Min(baseName.Length, 31 - marker.Length)] + marker;
        }

        return candidate;
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
            return EditorText.Get(EditorString.Sheet) + "1";
        }

        string result = name.Trim();
        foreach (char bad in new[] { '\\', '/', '?', '*', '[', ']', ':' })
        {
            result = result.Replace(bad, ' ');
        }

        result = result.Trim();
        if (result.Length == 0)
        {
            return EditorText.Get(EditorString.Sheet) + "1";
        }

        return result.Length > 31 ? result[..31] : result;
    }

    // Maps a cell's horizontal alignment, defaulting General to Excel's behavior
    // (numbers right-aligned, text left-aligned).
    private static TextAlignment AlignmentOf(XLAlignmentHorizontalValues horizontal, XLDataType dataType)
    {
        return horizontal switch
        {
            XLAlignmentHorizontalValues.Center or XLAlignmentHorizontalValues.CenterContinuous => TextAlignment.Center,
            XLAlignmentHorizontalValues.Right => TextAlignment.Right,
            XLAlignmentHorizontalValues.Left => TextAlignment.Left,
            _ => dataType == XLDataType.Number ? TextAlignment.Right : TextAlignment.Left,
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
