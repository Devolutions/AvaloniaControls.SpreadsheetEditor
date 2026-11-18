namespace SpreadsheetEditor.Controls.Xlsx;

using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

using global::SpreadsheetEditor.Controls.Models;

public static partial class XlsxFile
{
    private static string WorksheetRelationshipsPartPath(string part)
    {
        int slash = part.LastIndexOf('/');
        return part[..(slash + 1)] + "_rels/" + part[(slash + 1)..] + ".rels";
    }

    private static IReadOnlyDictionary<string, string> ReadHyperlinkRelationships(ZipArchive archive, string part)
    {
        string relationshipsPart = WorksheetRelationshipsPartPath(part);
        if (archive.GetEntry(relationshipsPart) is null) return new Dictionary<string, string>();
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        return ReadXmlPart(archive, relationshipsPart).Root!.Elements(packageRelationships + "Relationship")
            .Where(element => element.Attribute("Type")?.Value.EndsWith("/hyperlink", System.StringComparison.Ordinal) == true &&
                element.Attribute("Id") is not null && element.Attribute("Target") is not null)
            .ToDictionary(element => element.Attribute("Id")!.Value, element => element.Attribute("Target")!.Value);
    }

    // Excel accepts link targets that System.Uri rejects. Keep them opaque through SDK serialization
    // so relative paths and malformed URLs remain editable without blocking an unrelated cell save.
    private static void RestoreHyperlinks(MemoryStream package,
        IReadOnlyDictionary<int, Dictionary<CellAddress, string>> targets)
    {
        if (targets.Count == 0) return;
        package.Position = 0;
        using ZipArchive archive = new(package, ZipArchiveMode.Update, leaveOpen: true);
        XNamespace officeRelationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        XNamespace packageRelationships = "http://schemas.openxmlformats.org/package/2006/relationships";
        foreach ((int sheetIndex, Dictionary<CellAddress, string> sheetTargets) in targets)
        {
            string part = WorksheetPartPath(archive, sheetIndex);
            string relationshipsPart = WorksheetRelationshipsPartPath(part);
            XDocument worksheet = ReadXmlPart(archive, part);
            XDocument relationships = ReadXmlPart(archive, relationshipsPart);
            Dictionary<string, XElement> byId = relationships.Root!.Elements(packageRelationships + "Relationship")
                .ToDictionary(element => element.Attribute("Id")!.Value);
            int restored = 0;
            foreach (XElement hyperlink in worksheet.Descendants(Ssml + "hyperlink"))
            {
                if (!TryParseCellAddress(hyperlink.Attribute("ref")?.Value, out int row, out int column) ||
                    !sheetTargets.TryGetValue(new CellAddress(row, column), out string? target)) continue;
                string id = hyperlink.Attribute(officeRelationships + "id")!.Value;
                byId[id].SetAttributeValue("Target", target);
                restored++;
            }
            if (restored != sheetTargets.Count)
            {
                throw new InvalidDataException("A worksheet hyperlink relationship is missing.");
            }
            archive.GetEntry(relationshipsPart)!.Delete();
            using Stream output = archive.CreateEntry(relationshipsPart, CompressionLevel.Optimal).Open();
            using XmlWriter writer = XmlWriter.Create(output, new XmlWriterSettings
            {
                Encoding = new UTF8Encoding(false), NewLineHandling = NewLineHandling.Entitize,
            });
            relationships.Save(writer);
        }
    }
}
