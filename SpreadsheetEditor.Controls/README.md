# Devolutions.AvaloniaControls.SpreadsheetEditor

An Excel-like spreadsheet editor control for [Avalonia UI](https://avaloniaui.net/). It uses a custom
grid (not the Avalonia `DataGrid`) so cells can merge and host rich content, and it ships with:

- A custom rendering grid with cell merging, frozen headers, and drag-to-select.
- Rich in-cell editing (bold/italic/underline/strikethrough, fonts, colors, alignment).
- A hand-rolled `=` formula engine (arithmetic, ranges, `SUM/AVERAGE/IF/…`, named ranges, live recalc).
- Charts, multi-sheet tabs, infinite rows/columns, undo/redo, clipboard, and HTML print.

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
