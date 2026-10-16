namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Linq;

using ClosedXML.Excel;

using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Models;

public static partial class XlsxFile
{
    private static CellValue? ReadFormulaValue(XLCellValue value) => value.Type switch
    {
        XLDataType.Number => CellValue.FromNumber(value.GetNumber()),
        XLDataType.Boolean => CellValue.FromBool(value.GetBoolean()),
        XLDataType.Text => CellValue.FromText(DecodeCachedText(value.GetText())),
        XLDataType.Error => CellValue.FromError(CachedFormulaText(value, numberFormat: null)),
        XLDataType.DateTime => CellValue.FromNumber(value.GetDateTime().ToOADate()),
        XLDataType.TimeSpan => CellValue.FromNumber(value.GetTimeSpan().TotalDays),
        _ => null,
    };

    private static CellValue? ReadPreviewFormulaValue(XElement cell, IReadOnlyDictionary<int, string> sharedStrings)
    {
        if (cell.Element(Ssml + "f") is null || cell.Element(Ssml + "v") is not { } cached)
        {
            return null;
        }
        string raw = cached.Value;
        return (cell.Attribute("t")?.Value ?? string.Empty) switch
        {
            "b" => CellValue.FromBool(raw == "1"),
            "e" => CellValue.FromError(raw),
            "str" => CellValue.FromText(DecodeCachedText(raw)),
            "s" when int.TryParse(raw, out int index) && sharedStrings.TryGetValue(index, out string? text) => CellValue.FromText(text),
            "" or "n" when double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) => CellValue.FromNumber(number),
            _ => null,
        };
    }

    private static CellValue? SavedFormulaValue(CellData cell) => cell.FormulaValue ??
        (cell.FormulaNumericValue is { } number ? CellValue.FromNumber(number) :
            cell.FormulaResult is { } text ? CellValue.FromText(text) : null);

    // ClosedXML doesn't expose a setter for a formula's cached value. Persist the already-known
    // results into OOXML after serialization, without evaluating formulas in a second engine.
    private static void RestoreFormulaCaches(MemoryStream package, IReadOnlyList<XlsxSheetWriteData> sheets)
    {
        package.Position = 0;
        using ZipArchive archive = new(package, ZipArchiveMode.Update, leaveOpen: true);
        XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        Dictionary<string, string> targets = ReadXmlPart(archive, "xl/_rels/workbook.xml.rels")
            .Descendants(packageRelationships + "Relationship")
            .Where(element => element.Attribute("Id") is not null && element.Attribute("Target") is not null)
            .ToDictionary(element => element.Attribute("Id")!.Value,
                element => ResolvePartPath("xl/workbook.xml", element.Attribute("Target")!.Value));
        string[] parts = ReadXmlPart(archive, "xl/workbook.xml").Descendants(Ssml + "sheet")
            .Select(sheet => targets[sheet.Attribute(relationships + "id")!.Value]).ToArray();

        for (int index = 0; index < sheets.Count; index++)
        {
            Dictionary<string, CellValue> cached = new(StringComparer.OrdinalIgnoreCase);
            IReadOnlyList<IReadOnlyList<CellData>> rows = sheets[index].Data.Cells;
            for (int row = 0; row < rows.Count; row++)
            {
                for (int column = 0; column < rows[row].Count; column++)
                {
                    CellData cell = rows[row][column];
                    if (FormulaEngine.IsFormula(cell.Text) && SavedFormulaValue(cell) is { } value)
                    {
                        cached[new CellAddress(row, column).ToString()] = value;
                    }
                }
            }
            if (cached.Count == 0)
            {
                continue;
            }

            ZipArchiveEntry entry = archive.GetEntry(parts[index])!;
            using MemoryStream rewritten = new();
            using (Stream input = entry.Open())
            using (XmlReader reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit }))
            using (XmlWriter writer = XmlWriter.Create(rewritten, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                NewLineHandling = NewLineHandling.Entitize,
            }))
            {
                reader.Read();
                while (!reader.EOF)
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == Ssml.NamespaceName &&
                        reader.LocalName is "worksheet" or "sheetData" or "row")
                    {
                        // Keep the large worksheet streaming; only a cached formula cell becomes an XElement.
                        writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
                        writer.WriteAttributes(reader, defattr: true);
                        if (reader.IsEmptyElement) writer.WriteEndElement();
                        reader.Read();
                    }
                    else if (reader.NodeType == XmlNodeType.EndElement)
                    {
                        writer.WriteFullEndElement();
                        reader.Read();
                    }
                    else if (reader.NodeType == XmlNodeType.Element && reader.NamespaceURI == Ssml.NamespaceName &&
                             reader.LocalName == "c" && cached.TryGetValue(reader.GetAttribute("r") ?? string.Empty, out CellValue value))
                    {
                        XElement cell = (XElement)XNode.ReadFrom(reader);
                        if (cell.Element(Ssml + "f") is { } formula)
                        {
                            (string type, string text) = SerializedFormulaValue(value);
                            if (type == "str") text = EncodeCachedText(text);
                            cell.SetAttributeValue("t", type);
                            cell.Elements(Ssml + "v").Remove();
                            cell.Elements(Ssml + "is").Remove();
                            formula.AddAfterSelf(new XElement(Ssml + "v", text));
                        }
                        cell.WriteTo(writer);
                    }
                    else
                    {
                        writer.WriteNode(reader, defattr: true);
                    }
                }
            }
            entry.Delete();
            using Stream output = archive.CreateEntry(parts[index], CompressionLevel.Optimal).Open();
            rewritten.Position = 0;
            rewritten.CopyTo(output);
        }
    }

    private static (string Type, string Text) SerializedFormulaValue(CellValue value) => value.Kind switch
    {
        CellValueKind.Number when double.IsNaN(value.Number) => ("e", FormulaErrors.Num),
        CellValueKind.Number when double.IsInfinity(value.Number) => ("e", FormulaErrors.Div0),
        CellValueKind.Number => ("n", value.Number.ToString("R", CultureInfo.InvariantCulture)),
        CellValueKind.Boolean => ("b", value.Boolean ? "1" : "0"),
        CellValueKind.Error when value.Text != FormulaErrors.Circular => ("e", value.Text),
        _ => ("str", value.AsText()),
    };

    // Excel escapes control characters and literal escape-looking text independently of XML entities.
    private static string EncodeCachedText(string text) => Regex.Replace(text,
        @"_x[0-9A-Fa-f]{4}_|[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]",
        static match => match.Length > 1 ? "_x005F_" + match.Value[1..] :
            "_x" + ((int)match.Value[0]).ToString("X4", CultureInfo.InvariantCulture) + "_");

    private static string DecodeCachedText(string text) => Regex.Replace(text, @"_x([0-9A-Fa-f]{4})_",
        static match => ((char)int.Parse(match.Groups[1].Value, NumberStyles.HexNumber, CultureInfo.InvariantCulture)).ToString());
}
