# Plan — `SpreadsheetEditor` reusable Avalonia component

## 1. Overview

Build a **reusable spreadsheet editor control** for Avalonia, exposed as a standalone
control library and consumed by the existing `SpreadSheetEditor` desktop app as a demo
host. The control is named **`SpreadsheetEditor`** and renders an editable grid of cells
(rows × columns) with selection, keyboard navigation, inline editing, copy/paste, and
row/column add/remove/resize.

### Decisions locked in
| Area | Choice |
|------|--------|
| Rendering | Built on Avalonia **DataGrid** (`Avalonia.Controls.DataGrid`) |
| Scope (v1) | **Core editing + rich cell formatting** — no formulas, no file import/export yet |
| Cell editing | Avalonia **`RichTextEditor`** (Pro) for the active cell; rich per-cell content |
| Packaging | **Separate control library** (`SpreadsheetEditor.Controls`), NuGet-ready |
| Patterns | **MVVM** via `CommunityToolkit.Mvvm` |
| Target | `net10.0`, Avalonia `12.0.2`, Fluent theme |
| Licensing | Avalonia **Pro** — `AvaloniaUILicenseKey` required at build |

### Cell formatting via Avalonia RichTextEditor (Pro)
Use the official **`Avalonia.Controls.RichTextEditor`** control (Avalonia Pro tier) for cell
content and formatting. It provides a `FlowDocument` model (`RichRun`, `RichBold`,
`RichItalic`, `RichUnderline`, `Paragraph`…), `TextRange`/`TextPointer` for selection,
`range.ApplyPropertyValue(RichTextElement.ForegroundProperty, …)` for programmatic
formatting, a built-in `UndoManager`, and `IDocumentSerializer` round-tripping (RTF / DOCX /
XAML / plain text). Docs: <https://docs.avaloniaui.net/controls/input/text-input/richtexteditor>.

**Performance pattern (important):** do **not** place a live `RichTextEditor` in every
DataGrid cell — that would defeat row virtualization. Instead:
- **Active cell only** hosts a live `RichTextEditor` instance — surfaced in the formula bar
  and used as the in-cell edit template when editing. There is at most one live editor.
- **Display cells** render lightweight read-only formatted text (a `TextBlock` with inlines
  derived from the cell's document, or a non-interactive lightweight renderer). Validate the
  realized-cell cost during the M1 spike.

### Licensing — `AvaloniaUILicenseKey`
The `RichTextEditor` requires a Pro license key at build time, added as an
`<AvaloniaUILicenseKey Include="…"/>` item. Place it in a shared **`Directory.Build.props`**
at the repo root (or each csproj that pulls in the Pro control) so library + sample app both
build. **Security note:** the key is a credential — it (and this repo) should be treated
accordingly. Decide deliberately whether to commit it or inject it via CI secret /
`Directory.Build.props.user`; if committed to a public repo it is effectively exposed.

## 2. Goals / Non-goals

**Goals (v1)**
- Drop-in `SpreadsheetEditor` control reusable across Avalonia apps.
- A1-style addressing (columns `A, B, C…`, rows `1, 2, 3…`).
- **Excel-style headers:** a lettered column header row (`A, B, C…`) and a numbered row
  header column (`1, 2, 3…`), plus a top-left "select-all" corner. Clicking a column or
  row header selects the entire column/row; the active cell's column/row headers highlight.
- **Cell editor text box (formula bar):** a text box above the grid that displays and edits
  the active cell's content, two-way synced with the grid; preceded by a Name Box showing
  the active cell address (e.g. `B3`). Editing in either the bar or the cell stays in sync.
- Cell selection (single + range), keyboard navigation (arrows, Tab, Enter, Home/End).
- Inline cell editing (F2 / double-click / type-to-edit), commit on Enter/Tab, cancel on Esc.
- Clipboard: copy / cut / paste of single cells and rectangular ranges (TSV format).
- Add / remove / resize rows and columns.
- **Rich cell formatting** via Avalonia `RichTextEditor`: background, foreground, bold /
  italic / underline, alignment — and mixed-run rich text *within* a cell — applied to the
  selection via a toolbar.
- A **formatting toolbar** above the editor with these controls (see layout in §4).
- Clean, observable data model decoupled from the view so host apps can bind and inspect data.

**Non-goals (v1 — deferred, noted for forward-compatibility)**
- Formula engine / recalculation graph.
- File import/export (CSV/XLSX) **as a library feature** (the sample app does open/close —
  see §4; it can leverage the RichTextEditor serializers for rich round-trips).
- Number formats, merged cells, frozen panes.
- Undo/redo across the *sheet* (the RichTextEditor brings its own per-cell `UndoManager`;
  a sheet-wide undo stack is deferred — design the model so it can be added).

## 3. Why DataGrid — and the key constraint

Avalonia's `DataGrid` is built for **typed rows with a fixed set of columns**, each column
bound to a *property* of the row item. A spreadsheet needs **dynamic, positional columns**
(A, B, C…) and free-form cells. We bridge this gap as follows:

- Set `AutoGenerateColumns="false"` and **generate columns in code** (one per spreadsheet
  column), each a `DataGridTemplateColumn` whose cell template binds to a positional cell.
- Each **row item exposes an indexer / observable cell collection** so a column at index `i`
  binds to `Cells[i]`. Binding path: `Cells[i].Value` (with `INotifyPropertyChanged`).
- Column add/remove/resize is driven by mutating the model + regenerating/patching the
  `DataGrid.Columns` collection via a behavior/attached helper.

**Risks to validate early (spike in Milestone 1):**
- Per-cell range selection across columns: DataGrid's native selection is row-oriented.
  We will likely implement a custom selection model layered over DataGrid cell visuals.
- `DataGridTemplateColumn` editing UX (begin/commit/cancel edit) vs. our keyboard model.
- Performance with many dynamic columns (DataGrid virtualizes rows well; column count
  should stay modest — document a recommended cap, e.g. a few hundred columns).

If the spike shows DataGrid cell-level selection is too constraining, the fallback is a
custom-drawn grid — but the model/API below is designed to be rendering-agnostic so that
swap would not break the public API.

## 4. Project structure

```
SpreadSheetEditor.sln
├─ SpreadsheetEditor.Controls/          # NEW reusable control library
│  ├─ SpreadsheetEditor.Controls.csproj # Avalonia control lib, net10.0
│  ├─ Controls/
│  │  ├─ SpreadsheetEditor.axaml(.cs)   # the templated control (toolbar + formula bar + grid)
│  │  ├─ FormattingToolbar.axaml(.cs)   # drives RichTextEditor.ApplyPropertyValue on selection
│  │  └─ CellEditorBar.axaml(.cs)       # Name Box + active-cell RichTextEditor (formula bar)
│  ├─ ViewModels/
│  │  ├─ SpreadsheetViewModel.cs
│  │  ├─ RowViewModel.cs
│  │  └─ CellViewModel.cs
│  ├─ Models/
│  │  ├─ SpreadsheetDocument.cs
│  │  ├─ Cell.cs                        # rich content (FlowDocument or serialized XAML/RTF)
│  │  └─ CellAddress.cs                 # A1 <-> (row,col) conversion
│  ├─ Selection/
│  │  └─ SelectionModel.cs              # active cell + range
│  ├─ Clipboard/
│  │  └─ TsvClipboard.cs                # serialize/parse rectangular ranges
│  ├─ Behaviors/
│  │  └─ DataGridColumnsBehavior.cs     # dynamic columns bridge
│  └─ Themes/
│     └─ Generic.axaml                  # default control template + styles
└─ SpreadSheetEditor/                   # existing app -> demo host
   ├─ MainWindow.axaml(.cs)             # hosts <se:SpreadsheetEditor/> + menu/toolbar
   └─ SpreadsheetFile.cs                # app-side open/save (TSV/CSV) for testing
```

- Library references: `Avalonia`, `Avalonia.Controls.DataGrid` (12.0.2),
  `Avalonia.Controls.RichTextEditor` (Pro), `CommunityToolkit.Mvvm`. It should **not**
  reference `Avalonia.Desktop` (host concern).
- App adds a `ProjectReference` to the library and registers the DataGrid theme include.
- `AvaloniaUILicenseKey` lives in a root `Directory.Build.props` so both projects build the
  Pro control (see Licensing note above). For the sample app's open/close, add the
  serializer package(s) needed (e.g. `Avalonia.Controls.Documents.Serialization.Rtf`).

### Sample app (open/close a spreadsheet)

The existing `SpreadSheetEditor` desktop project doubles as the **sample/test host**. It
exercises the control and provides a real document lifecycle:

- **Menu / toolbar:** `New`, `Open…`, `Save`, `Save As…`, `Close`.
- **Open** — uses Avalonia's `IStorageProvider` file picker to choose a file, parses it into
  the control's `Document`, and binds it. **Close** — clears the control back to an empty
  sheet, prompting if there are unsaved changes (track a dirty flag via `CellValueChanged`).
  The window title reflects the current file name + dirty marker.
- **Formats:**
  - `.tsv`/`.csv` — value-only, via `SpreadsheetFile.cs` (built).
  - `.xlsx` — **value + whole-cell formatting** (bold/italic/underline, font + fill color),
    via `XlsxFile.cs` backed by **ClosedXML** (MIT). Reads the first worksheet, writes a
    single "Sheet1"; pickers route by extension; carries `CellData` to/from the control (built).
  - *(Planned)* mixed-run rich text within a cell via the RichTextEditor `IDocumentSerializer`
    (XAML/RTF). Current fidelity is whole-cell (commit reads the first run).
- File I/O lives in the **app** (`SpreadsheetFile.cs`, `XlsxFile.cs`), not the library. This
  keeps v1's "no library file I/O" scope while still giving a testable open/close flow.
- The picker requires the desktop lifetime + a top-level window, which is exactly why this
  belongs in the host app rather than the reusable control.

**Control layout** (vertical stack inside the control template):

```
┌───────────────────────────────────────────────┐
│ B  I  U │ 🅰▾ 🖌▾ │ ⯇ ⯈ ☰ │ + - rows/cols       │  ← FormattingToolbar
├───────────────────────────────────────────────┤
│ [ B3 ▼ ]  fx │ =cell editor text box────────── │  ← CellEditorBar (Name Box + editor)
├───────────────────────────────────────────────┤
│   │ A │ B │ C │ D │ …                           │  ← column headers (A,B,C…)
│ 1 │   │   │   │   │                             │
│ 2 │   │   │   │   │    grid (DataGrid)          │  ← row headers (1,2,3…) + cells
│ 3 │   │ ▣ │   │   │                             │
└───────────────────────────────────────────────┘
  ↑ select-all corner   ↑ active cell ▣
```

- **`FormattingToolbar`** sits at the top. Toggle buttons (bold/italic/underline) and color
  pickers (background/foreground) and alignment buttons act on the active cell's
  `RichTextEditor` via `TextRange.ApplyPropertyValue(...)` over the current selection;
  button state reflects the editor's current selection. A `ShowToolbar` property hides it
  for embedders who want a chromeless grid.

- **Column headers** come from per-column `Header` metadata (`A, B, C…`); the DataGrid's
  built-in **`RowHeader`** renders the row numbers (`1, 2, 3…`). The top-left corner cell
  is a select-all affordance. Header click selects the whole column/row; the active cell's
  row/column headers are visually highlighted (Excel behavior).
- The **`CellEditorBar`** hosts the single live `RichTextEditor`, bound to the active cell's
  rich content. It IS the in-cell editor surface for the active cell, so the bar and the cell
  never diverge. (Optionally the editor can be reparented into the active grid cell while
  editing; simplest first cut keeps it in the bar.)

## 5. Data model (rendering-agnostic core)

- **`CellAddress`** — value type; converts between `(row, col)` and A1 (`"B3"`). Pure, unit-tested.
- **`Cell`** — rich content: holds a `FlowDocument` (or its serialized XAML/RTF form for
  storage) plus a cached plain-text projection for display/TSV/search. Empty cells stay
  cheap (null document = empty). Rows hold cells sparsely or densely (decide in M1).
- **`SpreadsheetDocument`** — the source of truth: dimensions (rows/cols), cell access by
  address, and mutation ops (set content, apply formatting to a cell/range, insert/delete
  row/col). Raises change notifications.
- **ViewModels** wrap the model for binding:
  - `CellViewModel` : `ObservableObject` with the cell's document, a plain-text projection,
    `IsSelected`, `IsActive`. Display cells bind a lightweight read-only rendering of the
    document; only the active cell is backed by the live `RichTextEditor`.
  - `RowViewModel` : exposes `Cells` as an indexable observable collection (the DataGrid binds `Cells[i]`).
  - `SpreadsheetViewModel` : `ObservableCollection<RowViewModel> Rows`, column metadata
    (`ColumnViewModel` with width + header), selection, and `[RelayCommand]`s.

Design the document mutations as discrete operations so **undo/redo** can be layered later
without reworking the API.

## 6. Public API (control surface for host apps)

The control should be usable both code-first and XAML-first. Proposed surface:

```csharp
// XAML
<se:SpreadsheetEditor x:Name="Sheet"
                      RowCount="100" ColumnCount="26"
                      ShowHeaders="True" ShowFormulaBar="True" ShowToolbar="True" />
```

- **Styled/Direct properties:** `RowCount`, `ColumnCount`, `ShowHeaders` (column letters +
  row numbers), `ShowFormulaBar` (the cell editor text box / Name Box), `ShowToolbar`
  (formatting toolbar), `DefaultColumnWidth`, `DefaultRowHeight`, `IsReadOnly`,
  `SelectedRange`, `ActiveCell`.
- **Methods:** `GetCell(address)`, `SetCell(address, value)` (plain text), `Clear()`,
  `InsertRow/Column`, `DeleteRow/Column`, `CopySelection()`, `Paste()`,
  `ApplyFormat(range, property, value)` / `ClearFormat(range)` (wraps
  `TextRange.ApplyPropertyValue` over the affected cells' documents).
- **Events:** `CellValueChanged`, `SelectionChanged`, `ActiveCellChanged`.
- **Data access:** a `Document` property (or `LoadData(IEnumerable<...>)`) so host apps can
  bind/serialize. Keep the public API free of DataGrid types so the renderer can change.

## 7. Milestones

### M0 — Library scaffolding + license
- Create `SpreadsheetEditor.Controls` class library; add to solution.
- Add package refs (`Avalonia.Controls.DataGrid`, `Avalonia.Controls.RichTextEditor`,
  `CommunityToolkit.Mvvm`).
- Add `AvaloniaUILicenseKey` to a root `Directory.Build.props`; confirm the Pro control
  activates (a trivial `RichTextEditor` in the demo renders without a license watermark/error).
- Add `Themes/Generic.axaml`, wire the Fluent + DataGrid theme includes in the demo app.
- Demo host renders an empty `SpreadsheetEditor`. Exit: project builds, runs, Pro control activates.

### M1 — Dynamic-column DataGrid spike + rich-cell rendering (critical risk)
- Implement `CellAddress`, `Cell`, `SpreadsheetDocument`, ViewModels.
- Build `DataGridColumnsBehavior` to generate N columns bound to `Cells[i]`.
- Render a read-only grid of placeholder data with **Excel-style headers**: lettered
  column headers (`A, B, C…`), DataGrid `RowHeader` row numbers (`1, 2, 3…`), and the
  top-left select-all corner.
- **Validate the constraints in §3** including the rich-text rendering pattern: display
  cells use a lightweight read-only rendering of the cell document (NOT a live editor per
  cell); measure realized-cell cost while scrolling. Confirm one shared `RichTextEditor` for
  the active cell is viable. Adjust before building further.
- Exit: grid displays bound rich content with correct A1 headers at acceptable scroll perf.

### M2 — Selection, keyboard navigation & cell editor bar
- `SelectionModel`: active cell + rectangular range; visual highlight via cell styles,
  including highlighting the active cell's row/column headers.
- Arrow keys, Tab/Shift+Tab, Enter/Shift+Enter, Home/End, Ctrl+Arrow, Shift+select range.
- Header clicks select the whole column/row.
- **`CellEditorBar`**: Name Box shows the active cell address; hosts the single live
  `RichTextEditor` bound to the active cell's document (commit on Enter, cancel on Esc).
  This is the editing surface reused in M3.
- `SelectionChanged` / `ActiveCellChanged` events. Exit: full keyboard nav, range select,
  and a working formula bar (RichTextEditor) synced to the active cell.

### M3 — Inline editing
- Enter edit mode (F2, double-click, type-to-replace); commit (Enter/Tab), cancel (Esc).
- In-cell editing uses the same active-cell `RichTextEditor` (reparented into the cell, or
  the simplest first cut: edit via the formula bar with the cell mirroring live).
- `CellValueChanged` event; write-through to the document. Honor `IsReadOnly`.
- Exit: editing works via keyboard and mouse with correct commit/cancel semantics.

### M4 — Clipboard
- `TsvClipboard`: copy/cut/paste rectangular ranges as TSV (Excel-compatible).
- Wire Ctrl+C/X/V; paste grows the sheet if needed. Exit: round-trips with Excel/Sheets.

### M5 — Structure editing
- Insert/delete rows & columns (with address shifting); resize rows/columns via headers.
- Exit: add/remove/resize work and selection stays consistent.

### M6 — Formatting toolbar (RichTextEditor)
- `FormattingToolbar`: bold/italic/underline toggles, background + foreground color pickers,
  alignment buttons — each calls `TextRange.ApplyPropertyValue(RichTextElement.*Property, …)`
  on the active editor's selection.
- Toolbar button state reflects the editor's current selection (queried back from the range).
- Range formatting across multiple cells applies the property to each affected cell's document.
- Exit: select text/range, apply bg/fg/bold/italic/underline/alignment; toolbar tracks the
  active selection; per-cell undo works via the editor's `UndoManager`.

### M7 — Sample app: open/close + polish + docs
- **Sample app document lifecycle:** New / Open… / Save / Save As… / Close via menu/toolbar,
  using `IStorageProvider` pickers and `SpreadsheetFile.cs` (TSV/CSV, reusing M4's
  `TsvClipboard`). Dirty-flag tracking, unsaved-changes prompt on Close, title-bar file name.
- Demo app exercises every feature; README with usage + the public API.
- Mark NuGet metadata on the library csproj. Exit: open a file, edit, save, close — round-trips
  cleanly; library ready to reference from other projects.

## 8. Testing strategy
- **Unit tests** (new `SpreadsheetEditor.Tests` xUnit project): `CellAddress` conversions,
  `SpreadsheetDocument` mutations + address shifting, `SelectionModel` math, `TsvClipboard`
  parse/serialize round-trips. These are the highest-value, view-independent tests.
- **Manual/headless UI checks** for editing & navigation via the demo app per milestone.
- Run `dotnet build` and `dotnet test` from the solution root after each milestone.

## 9. Open questions / things to confirm during M1
- Confirm `Avalonia.Controls.DataGrid` + `Avalonia.Controls.RichTextEditor` package versions
  compatible with Avalonia 12.0.2, and that the Pro control activates with the provided key.
- **Rich-cell rendering cost:** confirm the lightweight read-only document rendering for
  display cells is cheap enough while scrolling many rows; if not, fall back to plain-text
  display with formatting shown only in the active editor.
- Cell storage: dense 2D vs. sparse dictionary (affects memory for large/empty sheets);
  store `FlowDocument` vs. serialized XAML/RTF for inactive cells.
- Whether per-cell range selection is acceptable on DataGrid or warrants the custom-grid fallback.
- Reparenting the single live `RichTextEditor` into the active cell vs. editing via the formula bar.
- Recommended practical limits (max rows/cols) to document for consumers.

## 10. Sequencing summary
M0 scaffold + license → **M1 spike (de-risk DataGrid + rich-cell rendering)** → M2 selection +
formula bar (RichTextEditor) → M3 editing → M4 clipboard → M5 structure → M6 formatting
toolbar → M7 sample app open/close + polish. The DataGrid + rich-rendering risks are
front-loaded in M1 so any pivot happens before the API-dependent work is built.
