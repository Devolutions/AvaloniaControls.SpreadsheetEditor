namespace SpreadsheetEditor.Controls.Xlsx;

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;

using ClosedXML.Excel;

using global::SpreadsheetEditor.Controls.Models;

public static partial class XlsxFile
{
    // Read only the worksheet's relationships and small table parts, never scan unrelated sheets
    // or materialize their cell data just to find the first-screen header dropdowns.
    private static List<SheetTable> ReadTableDefinitions(ZipArchive archive, string worksheetPart,
        CancellationToken cancellationToken, List<SheetColumnFilter> filters)
    {
        List<SheetTable> result = new();
        int slash = worksheetPart.LastIndexOf('/');
        string relationshipsPart = worksheetPart[..(slash + 1)] + "_rels/" + worksheetPart[(slash + 1)..] + ".rels";
        if (archive.GetEntry(relationshipsPart) is not { } relationshipsEntry) return result;
        using Stream stream = relationshipsEntry.Open();
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
        XElement relationships = XElement.Load(reader);
        foreach (XElement relationship in relationships.Elements())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (relationship.Attribute("Type")?.Value.EndsWith("/table", StringComparison.Ordinal) != true ||
                relationship.Attribute("TargetMode")?.Value == "External" || relationship.Attribute("Target")?.Value is not { } target)
                continue;
            string part = ResolvePartPath(worksheetPart, target);
            if (archive.GetEntry(part) is not { } entry) continue;
            using Stream tableStream = entry.Open();
            using XmlReader tableReader = XmlReader.Create(tableStream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit });
            XElement table = XElement.Load(tableReader);
            if (!CellRange.TryParse(table.Attribute("ref")?.Value ?? string.Empty, out CellRange range)) continue;
            XElement? autoFilter = table.Element(Ssml + "autoFilter");
            CellRange? filterRange = CellRange.TryParse(autoFilter?.Attribute("ref")?.Value ?? string.Empty, out CellRange parsed)
                ? parsed : null;
            XElement? style = table.Element(Ssml + "tableStyleInfo");
            static bool Flag(XElement? element, string name) => element?.Attribute(name)?.Value is "1" or "true";
            result.Add(new SheetTable
            {
                Name = table.Attribute("displayName")?.Value ?? table.Attribute("name")?.Value ?? "Table" + (result.Count + 1),
                Range = range,
                ColumnNames = table.Element(Ssml + "tableColumns")?.Elements(Ssml + "tableColumn")
                    .Select(column => column.Attribute("name")?.Value ?? string.Empty).ToArray() ?? Array.Empty<string>(),
                AutoFilterRange = filterRange,
                ShowHeaderRow = table.Attribute("headerRowCount")?.Value != "0",
                ShowTotalsRow = int.TryParse(table.Attribute("totalsRowCount")?.Value, out int totals) && totals > 0,
                StyleName = style?.Attribute("name")?.Value,
                ShowRowStripes = Flag(style, "showRowStripes"),
                ShowColumnStripes = Flag(style, "showColumnStripes"),
                EmphasizeFirstColumn = Flag(style, "showFirstColumn"),
                EmphasizeLastColumn = Flag(style, "showLastColumn"),
            });
            if (filterRange is { } filter) ReadValueFilters(autoFilter, filter, filters);
        }
        return result;
    }

    private static void ReadValueFilters(XElement? autoFilter, CellRange range, List<SheetColumnFilter> result)
    {
        if (autoFilter is null) return;
        foreach (XElement column in autoFilter.Elements(Ssml + "filterColumn"))
        {
            if (!int.TryParse(column.Attribute("colId")?.Value, out int offset) || offset < 0 || offset >= range.ColumnSpan ||
                column.Element(Ssml + "filters") is not { } filters) continue;
            List<string> values = filters.Elements(Ssml + "filter").Select(value => value.Attribute("val")?.Value ?? string.Empty).ToList();
            if (filters.Attribute("blank")?.Value is "1" or "true") values.Add(string.Empty);
            if (values.Count > 0) result.Add(new SheetColumnFilter(range.Row, range.Column + offset, values));
        }
    }

    private static void RestoreTables(MemoryStream package, IReadOnlyList<XlsxSheetWriteData> sheets)
    {
        if (!sheets.Any(sheet => sheet.Data.Tables.Count > 0)) return;
        package.Position = 0;
        using ZipArchive archive = new(package, ZipArchiveMode.Update, leaveOpen: true);
        XNamespace relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        XNamespace types = "http://schemas.openxmlformats.org/package/2006/content-types";
        XDocument contentTypes = ReadXmlPart(archive, "[Content_Types].xml");
        int id = 0;
        for (int index = 0; index < sheets.Count; index++)
        {
            SheetData data = sheets[index].Data;
            if (data.Tables.Count == 0) continue;
            string worksheetPart = WorksheetPartPath(archive, index);
            int slash = worksheetPart.LastIndexOf('/');
            string relsPart = worksheetPart[..(slash + 1)] + "_rels/" + worksheetPart[(slash + 1)..] + ".rels";
            XDocument rels = archive.GetEntry(relsPart) is null
                ? new XDocument(new XElement(packageRelationships + "Relationships")) : ReadXmlPart(archive, relsPart);
            XElement tableParts = new(Ssml + "tableParts");
            foreach (SheetTable definition in data.Tables)
            {
                CellRange range = definition.Range;
                if (range.RowSpan <= 0 || range.ColumnSpan <= 0) continue;
                int tableId = ++id;
                string tablePart = $"xl/tables/editorTable{tableId}.xml";
                string relationshipId = "rIdEditorTable" + tableId;
                while (rels.Root!.Elements().Any(element => element.Attribute("Id")?.Value == relationshipId)) relationshipId += "_";
                XElement table = new(Ssml + "table", new XAttribute("id", tableId),
                    new XAttribute("name", definition.Name), new XAttribute("displayName", definition.Name),
                    new XAttribute("ref", range.ToA1()), new XAttribute("headerRowCount", definition.ShowHeaderRow ? 1 : 0),
                    new XAttribute("totalsRowCount", definition.ShowTotalsRow ? 1 : 0),
                    new XAttribute("totalsRowShown", definition.ShowTotalsRow ? 1 : 0));
                if (definition.AutoFilterRange is { } filter)
                {
                    XElement autoFilter = new(Ssml + "autoFilter", new XAttribute("ref", filter.ToA1()));
                    foreach (SheetColumnFilter column in data.ColumnFilters.Where(column => column.HeaderRow == filter.Row &&
                        column.Column >= filter.Column && column.Column < filter.Column + filter.ColumnSpan))
                    {
                        XElement values = new(Ssml + "filters");
                        if (column.Values.Contains(string.Empty)) values.SetAttributeValue("blank", 1);
                        values.Add(column.Values.Where(value => value.Length > 0).Select(value => new XElement(Ssml + "filter", new XAttribute("val", value))));
                        autoFilter.Add(new XElement(Ssml + "filterColumn", new XAttribute("colId", column.Column - filter.Column), values));
                    }
                    table.Add(autoFilter);
                }
                XElement columns = new(Ssml + "tableColumns", new XAttribute("count", range.ColumnSpan));
                HashSet<string> names = new(StringComparer.OrdinalIgnoreCase);
                for (int column = 0; column < range.ColumnSpan; column++)
                {
                    string name = definition.ShowHeaderRow && range.Row < data.Cells.Count &&
                        range.Column + column < data.Cells[range.Row].Count ? data.Cells[range.Row][range.Column + column].Text :
                        column < definition.ColumnNames.Count ? definition.ColumnNames[column] : string.Empty;
                    if (string.IsNullOrWhiteSpace(name)) name = "Column" + (column + 1);
                    string unique = name;
                    for (int suffix = 2; !names.Add(unique); suffix++) unique = name + suffix;
                    columns.Add(new XElement(Ssml + "tableColumn", new XAttribute("id", column + 1), new XAttribute("name", unique)));
                }
                table.Add(columns);
                table.Add(new XElement(Ssml + "tableStyleInfo",
                    string.IsNullOrWhiteSpace(definition.StyleName) ? null : new XAttribute("name", definition.StyleName),
                    new XAttribute("showFirstColumn", definition.EmphasizeFirstColumn ? 1 : 0),
                    new XAttribute("showLastColumn", definition.EmphasizeLastColumn ? 1 : 0),
                    new XAttribute("showRowStripes", definition.ShowRowStripes ? 1 : 0),
                    new XAttribute("showColumnStripes", definition.ShowColumnStripes ? 1 : 0)));
                WriteTableXmlPart(archive, tablePart, new XDocument(table));
                rels.Root!.Add(new XElement(packageRelationships + "Relationship", new XAttribute("Id", relationshipId),
                    new XAttribute("Type", relationships.NamespaceName + "/table"), new XAttribute("Target", "../tables/editorTable" + tableId + ".xml")));
                tableParts.Add(new XElement(Ssml + "tablePart", new XAttribute(relationships + "id", relationshipId)));
                contentTypes.Root!.Add(new XElement(types + "Override", new XAttribute("PartName", "/" + tablePart),
                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")));
            }
            tableParts.SetAttributeValue("count", tableParts.Elements().Count());
            WriteTableXmlPart(archive, relsPart, rels);
            AppendTableParts(archive, worksheetPart, tableParts);
        }
        WriteTableXmlPart(archive, "[Content_Types].xml", contentTypes);
    }

    private static void WriteTableXmlPart(ZipArchive archive, string part, XDocument document)
    {
        archive.GetEntry(part)?.Delete();
        using Stream stream = archive.CreateEntry(part, CompressionLevel.Optimal).Open();
        document.Save(stream);
    }

    private static void AppendTableParts(ZipArchive archive, string part, XElement tableParts)
    {
        ZipArchiveEntry entry = archive.GetEntry(part)!;
        using MemoryStream rewritten = new();
        using (Stream stream = entry.Open())
        using (XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit }))
        using (XmlWriter writer = XmlWriter.Create(rewritten, new XmlWriterSettings
        { Encoding = new UTF8Encoding(false), NewLineHandling = NewLineHandling.Entitize }))
        {
            bool added = false;
            reader.Read();
            while (!reader.EOF)
            {
                if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "worksheet")
                {
                    writer.WriteStartElement(reader.Prefix, reader.LocalName, reader.NamespaceURI);
                    writer.WriteAttributes(reader, false);
                    reader.Read();
                }
                else if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "tableParts") reader.Skip();
                else if (!added && (reader.NodeType == XmlNodeType.EndElement && reader.LocalName == "worksheet" ||
                         reader.NodeType == XmlNodeType.Element && reader.LocalName == "extLst"))
                {
                    tableParts.WriteTo(writer);
                    added = true;
                }
                else writer.WriteNode(reader, false);
            }
        }
        entry.Delete();
        using Stream output = archive.CreateEntry(part, CompressionLevel.Optimal).Open();
        rewritten.Position = 0;
        rewritten.CopyTo(output);
    }

    private static void WriteValueFilters(IXLAutoFilter filter, CellRange? range, IReadOnlyList<SheetColumnFilter> columns)
    {
        if (range is not { } area || !filter.IsEnabled) return;
        foreach (SheetColumnFilter column in columns.Where(column => column.HeaderRow == area.Row &&
            column.Column >= area.Column && column.Column < area.Column + area.ColumnSpan))
        {
            foreach (string value in column.Values)
                filter.Column(column.Column - area.Column + 1).AddFilter(value, reapply: false);
        }
    }
}
