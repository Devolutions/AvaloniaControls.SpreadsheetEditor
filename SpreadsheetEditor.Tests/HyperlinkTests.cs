using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;

using Avalonia.Media;
using Avalonia.Headless.XUnit;

using ClosedXML.Excel;

using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class HyperlinkTests
{
    private static readonly XNamespace Main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public static TheoryData<string> Targets => new()
    {
        "https://devolutions.net/remote-desktop-manager/",
        "http://intranet portal/page.aspx",
        @"\\file server\share\report.xlsx",
        "file://srv 01/share/doc.txt",
        "mailto:support@example.com",
        "#'Links'!A1",
        "http://my_server_01/app",
        "http://[fe80::1%eth0]:99999/",
        @"..\docs\manual.pdf",
        "../docs/manual.pdf",
        "../docs/a%20file.pdf#page=3",
        "https://example.com/a%20file.pdf?value=a%2Fb&other=1#page=3",
    };

    [Theory]
    [MemberData(nameof(Targets))]
    public void SavingKeepsTheExactHyperlinkTarget(string target)
    {
        CellData cell = CellData.Plain("Linked value") with { Hyperlink = target, Foreground = Colors.Blue, Underline = true };
        using MemoryStream saved = new();
        XlsxFile.Write(saved, new[] { new[] { cell } });
        Assert.Equal(target, ReadSavedTarget(saved));
        saved.Position = 0;
        CellData reopened = XlsxFile.Read(saved).Cells[0][0];
        Assert.Equal(target, reopened.Hyperlink);
        Assert.Equal(cell.Text, reopened.Text);
        Assert.Equal(cell.Foreground, reopened.Foreground);
        Assert.True(reopened.Underline);
    }

    [Theory]
    [MemberData(nameof(Targets))]
    public void OpeningAnExistingWorkbookKeepsTheExactHyperlinkTarget(string target)
    {
        using MemoryStream original = CreateExistingWorkbook(target);
        CellData cell = XlsxFile.Read(original).Cells[0][0];
        Assert.Equal(target, cell.Hyperlink);
        Assert.Equal("Linked value", cell.Text);
    }

    [Fact]
    public void DistinctAndRepeatedLinksKeepTheirOwnRelationshipTargets()
    {
        string[] targets = { @"..\docs\manual.pdf", "http://intranet portal/page.aspx", "https://example.com/", @"..\docs\manual.pdf" };
        CellData[][] cells = { targets.Select(target => CellData.Plain("Link") with { Hyperlink = target }).ToArray() };
        using MemoryStream saved = new();
        XlsxFile.Write(saved, cells);
        saved.Position = 0;
        Assert.Equal(targets, XlsxFile.Read(saved).Cells[0].Select(cell => cell.Hyperlink));
    }

    [AvaloniaFact]
    public async Task AsyncSaveKeepsLinksAndEditsAlongsideUnopenedWorksheets()
    {
        string[] targets = { @"..\docs\manual.pdf", "http://intranet portal/page.aspx", "http://[fe80::1%eth0]:99999/" };
        using MemoryStream source = CreateExistingWorkbook(targets);
        EditorControl editor = new();
        try
        {
            await editor.OpenXlsxAsync(source);
            Assert.Equal(targets[0], editor.ViewModel.Rows[0].Cells[0].Hyperlink);
            object preview = editor.ViewModel;
            using MemoryStream untouched = new();
            await editor.SaveXlsxAsync(untouched);
            untouched.Position = 0;
            Assert.Equal(targets, XlsxFile.ReadWorkbook(untouched).Select(sheet => sheet.Data.Cells[0][0].Hyperlink));

            Stopwatch timeout = Stopwatch.StartNew();
            while (ReferenceEquals(preview, editor.ViewModel) && timeout.Elapsed < TimeSpan.FromSeconds(10)) await Task.Delay(10);
            Assert.NotSame(preview, editor.ViewModel);
            editor.ViewModel.SetCell(new CellAddress(0, 0), "Edited value");
            using MemoryStream saved = new();
            await editor.SaveXlsxAsync(saved);
            saved.Position = 0;
            var reopened = XlsxFile.ReadWorkbook(saved);
            Assert.Equal("Edited value", reopened[0].Data.Cells[0][0].Text);
            Assert.Equal(targets, reopened.Select(sheet => sheet.Data.Cells[0][0].Hyperlink));
        }
        finally
        {
            editor.NewSheet(1, 1);
        }
    }

    private static MemoryStream CreateExistingWorkbook(params string[] targets)
    {
        MemoryStream stream = new();
        using (XLWorkbook workbook = new())
        {
            for (int index = 0; index < targets.Length; index++)
            {
                IXLCell cell = workbook.AddWorksheet(index == 0 ? "Links" : "Links" + (index + 1)).Cell("A1");
                cell.Value = "Linked value";
                cell.SetHyperlink(new XLHyperlink(new Uri("https://example.com/placeholder")));
            }
            workbook.SaveAs(stream);
        }
        using (ZipArchive archive = new(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            for (int index = 0; index < targets.Length; index++)
            {
                string part = $"xl/worksheets/_rels/sheet{index + 1}.xml.rels";
                ZipArchiveEntry entry = archive.GetEntry(part)!;
                XDocument relationships;
                using (Stream input = entry.Open()) relationships = XDocument.Load(input);
                XElement hyperlink = Assert.Single(relationships.Root!.Elements(), element => element.Attribute("Type")!.Value.EndsWith("/hyperlink", StringComparison.Ordinal));
                hyperlink.SetAttributeValue("Target", targets[index]);
                entry.Delete();
                using Stream output = archive.CreateEntry(part).Open();
                relationships.Save(output);
            }
        }
        stream.Position = 0;
        return stream;
    }

    private static string ReadSavedTarget(MemoryStream stream)
    {
        stream.Position = 0;
        using ZipArchive archive = new(stream, ZipArchiveMode.Read, leaveOpen: true);
        using Stream worksheetStream = archive.GetEntry("xl/worksheets/sheet1.xml")!.Open();
        XElement hyperlink = Assert.Single(XDocument.Load(worksheetStream).Descendants(Main + "hyperlink"));
        if (hyperlink.Attribute("location") is { } location) return "#" + location.Value;
        using Stream relationshipsStream = archive.GetEntry("xl/worksheets/_rels/sheet1.xml.rels")!.Open();
        return XDocument.Load(relationshipsStream).Root!.Elements().Single(element =>
            element.Attribute("Id")!.Value == hyperlink.Attribute(Relationships + "id")!.Value).Attribute("Target")!.Value;
    }
}
