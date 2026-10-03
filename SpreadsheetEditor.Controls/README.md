# Devolutions.AvaloniaControls.SpreadsheetEditor

An Excel-like spreadsheet editor control for [Avalonia UI](https://avaloniaui.net/). It uses a custom
grid (not the Avalonia `DataGrid`) so cells can merge and host rich content, and it ships with:

- A custom rendering grid with cell merging, frozen headers, and drag-to-select.
- Rich in-cell editing (bold/italic/underline/strikethrough, fonts, colors, alignment).
- A hand-rolled `=` formula engine (arithmetic, ranges, `SUM/AVERAGE/IF/…`, named ranges, live recalc).
- Charts, multi-sheet tabs, infinite rows/columns, undo/redo, clipboard, and printing (cross-platform PDF export + HTML preview).

## Requirements

This control depends on the commercial **Avalonia Accelerate** packages
(`Avalonia.Controls.RichTextEditor`, `Avalonia.Controls.Charts`). You must supply your own
**Avalonia UI license key** (via the `<AvaloniaUILicenseKey>` MSBuild item or the standard licensing
mechanism) in the application that consumes this control. See <https://avaloniaui.net/>.

## Setup

Add the required theme includes to your `App.axaml` (alongside `FluentTheme`):

```xml
<Application.Styles>
    <FluentTheme />
    <StyleInclude Source="avares://Avalonia.Controls.RichTextEditor/Themes/Default.axaml" />
    <StyleInclude Source="avares://Avalonia.Controls.Charts/Themes/Default.axaml" />
</Application.Styles>
```

## Usage

```xml
xmlns:se="clr-namespace:SpreadsheetEditor.Controls;assembly=SpreadsheetEditor.Controls"
...
<se:SpreadsheetEditor x:Name="Editor" />
```

The control exposes an integration API similar to Avalonia's `RichTextEditor`: `IsReadOnly`,
`Undo()`/`Redo()` + `CanUndo`/`CanRedo`, `SelectAll()`, `SelectedRange`, and the `ContentChanged`,
`SelectionChanged`, and `IsReadOnlyChanged` events. Load data via `LoadSheet`, `LoadWorkbook`,
`LoadCellData`, or `LoadRows`; read it back with `GetCellData`/`GetRows`.

## Excel (.xlsx) files — included

Excel read/write is built in (backed by [ClosedXML](https://github.com/ClosedXML/ClosedXML) plus
OOXML chart parsing), so no extra package is needed:

```csharp
using System.IO;

// Load an .xlsx (every worksheet → one tab) into the control:
await using (FileStream fs = File.OpenRead("book.xlsx"))
    await editor.OpenXlsxAsync(fs); // keeps the UI responsive for large workbooks

// Save the complete workbook (every tab) back to .xlsx:
await using (FileStream fs = File.Create("out.xlsx"))
    editor.SaveXlsx(fs);
```

`OpenXlsx(Stream)` remains available for synchronous callers.

`OpenXlsxAsync` reads workbook tab names and colors immediately, then reads roughly one visible
viewport of the first worksheet, including its cached Excel formula results and charts. Other
worksheets remain lightweight, correctly styled placeholders until selected. Selecting a tab displays
its viewport-sized preview first and starts only that worksheet's full-fidelity import in the
background. The previous frame stays visible until that preview is ready, avoiding a placeholder
flash during the handoff. Preview sheets are read-only until their complete data is ready. A single
serialized ClosedXML reader is reused for the workbook, so browsing several tabs does not queue repeated
full-workbook parses. Opening a workbook preserves Excel's saved formula results instead of
recalculating every sheet. After a cell edit, any worksheet needed by a sheet-qualified reference such
as `Users!A1` loads on demand, then the affected formulas recalculate.

For headless / non-UI use, the readers/writers are also available directly via
`SpreadsheetEditor.Controls.Xlsx.XlsxFile` (`ReadWorkbook`/`Read`/`Write`) returning the library's
`SheetData` model.

## Printing (cross-platform)

Avalonia has no built-in print API, so the control offers two portable routes — both work on
Windows, macOS, and Linux:

```csharp
using System.IO;

// 1. PDF export (recommended). Self-contained, rendered with SkiaSharp — no browser needed.
//    Cell fills, borders, alignment, and rich run formatting are preserved; the sheet is scaled
//    to fit the page width and rows flow across pages.
await using (FileStream fs = File.Create("sheet.pdf"))
    editor.ExportPdf(fs);

editor.SavePdf("sheet.pdf");   // convenience overload for a file path
editor.PrintPdf();             // writes a temp PDF and sends it to the printer (see below)

// 2. HTML preview. Opens the sheet in the default browser, where the user prints (or saves a PDF)
//    via the browser's own print dialog.
string html = editor.ExportHtml();
editor.Print();
```

`PrintPdf()` sends the PDF straight to the printer: on Windows via the shell "print" verb (the
registered PDF handler prints to the default printer), and on Linux/macOS via `lpr` (the default CUPS
printer). `Print()` instead opens the HTML in the default browser (`ShellExecute`/`open`/`xdg-open`)
for browser-driven printing. Images are not drawn into the PDF/HTML (data only); charts are included
in the PDF.
