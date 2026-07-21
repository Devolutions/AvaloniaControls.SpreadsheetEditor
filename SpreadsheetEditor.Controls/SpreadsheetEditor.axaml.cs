namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.ViewModels;

// A reusable spreadsheet editor built on a custom SpreadsheetGrid (not the Avalonia
// DataGrid) so cells can span (merge) and host rich content. The grid renders each cell's text via
// inline runs (bold/italic/underline/colors); clicking a cell sets the active cell. The formula bar
// is a plain TextBox bound live to the active cell, and the toolbar applies whole-cell formatting.
public partial class SpreadsheetEditor : UserControl
{
    public static readonly StyledProperty<int> InitialRowCountProperty =
        AvaloniaProperty.Register<SpreadsheetEditor, int>(nameof(InitialRowCount), 50);

    public static readonly StyledProperty<int> InitialColumnCountProperty =
        AvaloniaProperty.Register<SpreadsheetEditor, int>(nameof(InitialColumnCount), 26);

    public static readonly StyledProperty<bool> IsReadOnlyProperty =
        AvaloniaProperty.Register<SpreadsheetEditor, bool>(nameof(IsReadOnly));

    private static readonly Color[] Palette =
    {
        Colors.Black, Colors.DimGray, Colors.Gray, Colors.White,
        Colors.Red, Colors.Orange, Colors.Gold, Colors.Yellow,
        Colors.Green, Colors.Teal, Colors.Blue, Colors.Indigo,
        Colors.Purple, Colors.HotPink, Colors.Brown, Colors.Transparent,
    };

    private const string DefaultFontName = "Calibri";

    private static readonly string[] FontNames =
    {
        "Calibri", "Cambria", "Arial", "Times New Roman", "Verdana", "Georgia",
        "Tahoma", "Segoe UI", "Courier New", "Consolas", "Comic Sans MS",
    };

    private static readonly int[] FontSizes = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72 };

    // Engine-supported functions grouped for the function library (name + signature).
    private static readonly (string Category, (string Name, string Signature)[] Functions)[] FunctionLibrary =
    {
        ("Math & Trig", new[]
        {
            ("SUM", "SUM(range)"), ("ABS", "ABS(number)"), ("SQRT", "SQRT(number)"),
            ("POWER", "POWER(number, power)"), ("MOD", "MOD(number, divisor)"), ("INT", "INT(number)"),
            ("ROUND", "ROUND(number, digits)"), ("ROUNDUP", "ROUNDUP(number, digits)"), ("ROUNDDOWN", "ROUNDDOWN(number, digits)"),
        }),
        ("Statistical", new[]
        {
            ("AVERAGE", "AVERAGE(range)"), ("MIN", "MIN(range)"), ("MAX", "MAX(range)"),
            ("COUNT", "COUNT(range)"), ("COUNTA", "COUNTA(range)"),
        }),
        ("Logical", new[]
        {
            ("IF", "IF(condition, if_true, if_false)"), ("IFERROR", "IFERROR(value, value_if_error)"),
            ("AND", "AND(a, b, …)"), ("OR", "OR(a, b, …)"), ("NOT", "NOT(value)"),
        }),
        ("Text", new[]
        {
            ("CONCAT", "CONCAT(text, …)"), ("CONCATENATE", "CONCATENATE(text, …)"), ("LEN", "LEN(text)"),
            ("LEFT", "LEFT(text, count)"), ("RIGHT", "RIGHT(text, count)"), ("MID", "MID(text, start, count)"),
            ("UPPER", "UPPER(text)"), ("LOWER", "LOWER(text)"), ("TRIM", "TRIM(text)"),
        }),
        ("Date & Time", new[]
        {
            ("TODAY", "TODAY()"), ("NOW", "NOW()"), ("YEAR", "YEAR(date)"),
            ("MONTH", "MONTH(date)"), ("DAY", "DAY(date)"), ("DATE", "DATE(year, month, day)"),
        }),
    };

    // The workbook: an ordered list of sheets, each its own view model; vm points at the active one.
    private readonly List<SpreadsheetViewModel> sheets = new();

    private SpreadsheetViewModel vm = new();

    private CellViewModel? activeCell;

    private TextBox? formulaBar;

    private bool suppressToolbar;

    private bool initialFocusDone;

    // The sheet whose UndoStateChanged/DataChanged we're currently subscribed to (for CanUndo + ContentChanged).
    private SpreadsheetViewModel? hookedSheet;

    public SpreadsheetEditor()
    {
        this.InitializeComponent();

        this.vm.Initialize(this.InitialRowCount, this.InitialColumnCount);
        this.vm.Name = "Sheet1";
        this.sheets.Add(this.vm);
        this.DataContext = this.vm;

        this.Grid.Sheet = this.vm;
        this.Grid.CellActivated += address => this.SetActiveCell(address);

        this.AddSheetButton.Click += (_, _) => this.AddSheet();

        this.CreateFormulaBar();

        this.BuildPalette(this.ForegroundPalette, this.ApplyForeground);
        this.BuildPalette(this.BackgroundPalette, this.ApplyBackground);

        // B/I/U + text color route to the in-cell editor (per-character when the rich editor is
        // active; whole-cell otherwise). Guarded so SyncToolbar setting IsChecked doesn't re-toggle.
        this.BoldButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar) { this.Grid.ToggleBold(); this.SyncFormatButtons(); } };
        this.ItalicButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar) { this.Grid.ToggleItalic(); this.SyncFormatButtons(); } };
        this.UnderlineButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar) { this.Grid.ToggleUnderline(); this.SyncFormatButtons(); } };
        this.StrikethroughButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar) { this.Grid.ToggleStrikethrough(); this.SyncFormatButtons(); } };

        // Keep B/I/U reflecting the caret/selection while editing.
        this.Grid.SelectionFormatChanged += (_, _) => this.SyncFormatButtons();

        this.SetupFontControls();
        this.AlignLeftButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar && this.AlignLeftButton.IsChecked == true) this.SetAlignment(TextAlignment.Left); };
        this.AlignCenterButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar && this.AlignCenterButton.IsChecked == true) this.SetAlignment(TextAlignment.Center); };
        this.AlignRightButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar && this.AlignRightButton.IsChecked == true) this.SetAlignment(TextAlignment.Right); };
        this.Grid.ChartEditRequested += async chart => await this.EditChartAsync(chart);

        // History + clipboard buttons.
        this.UndoButton.Click += (_, _) => this.Undo();
        this.RedoButton.Click += (_, _) => this.Redo();
        this.CutButton.Click += (_, _) => this.Grid.CutSelection();
        this.CopyButton.Click += (_, _) => this.Grid.CopySelection();
        this.PasteButton.Click += (_, _) => this.Grid.PasteSelection();
        this.PrintButton.Click += (_, _) => this.Print();

        // Spreadsheet-specific actions live in the overflow menu.
        this.OverflowButton.Flyout = this.BuildOverflowMenu();
        this.SymbolButton.Flyout = this.BuildSymbolFlyout();

        this.SetActiveCell(new CellAddress(0, 0));
        this.RebuildTabs();
        this.HookSheet();
        this.ApplyReadOnly();

        // Recolor the sheet tabs when the light/dark theme changes at runtime.
        this.ActualThemeVariantChanged += (_, _) => this.RebuildTabs();

        // Start with A1 selected (selection mode), not the auto-focused formula bar or an open editor.
        this.Loaded += (_, _) =>
        {
            this.RebuildTabs(); // theme is resolved by now — recolor tabs for the actual variant

            if (this.initialFocusDone)
            {
                return;
            }

            this.initialFocusDone = true;
            Dispatcher.UIThread.Post(this.Grid.FocusForSelection, DispatcherPriority.Background);
        };
    }

    // The current (active) sheet data (rendering-agnostic). Host apps read/write this.
    public SpreadsheetViewModel ViewModel => this.vm;

    // All sheets in the workbook, in tab order.
    public IReadOnlyList<SpreadsheetViewModel> Sheets => this.sheets;

    public int InitialRowCount
    {
        get => this.GetValue(InitialRowCountProperty);
        set => this.SetValue(InitialRowCountProperty, value);
    }

    public int InitialColumnCount
    {
        get => this.GetValue(InitialColumnCountProperty);
        set => this.SetValue(InitialColumnCountProperty, value);
    }

    // When true, the editor is view-only: no cell editing, typing, paste/cut/clear, or formatting.
    // Mirrors RichTextEditor.IsReadOnly so both controls integrate the same way.
    public bool IsReadOnly
    {
        get => this.GetValue(IsReadOnlyProperty);
        set => this.SetValue(IsReadOnlyProperty, value);
    }

    // The selected cell range (anchor↔focus), mirroring RichTextEditor's selection surface.
    public CellRange SelectedRange => this.Grid.SelectionRange;

    // True when there is a change to undo / redo on the active sheet.
    public bool CanUndo => this.vm.CanUndo;

    public bool CanRedo => this.vm.CanRedo;

    // Raised when cell content changes (after a recalculation) — analogous to RichTextEditor.ContentChanged.
    public event EventHandler? ContentChanged;

    // Raised when the active cell / selected range changes — analogous to RichTextEditor.SelectionChanged.
    public event EventHandler? SelectionChanged;

    // Raised when IsReadOnly changes — analogous to RichTextEditor.IsReadOnlyChanged.
    public event EventHandler? IsReadOnlyChanged;

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == IsReadOnlyProperty)
        {
            this.ApplyReadOnly();
        }
    }

    private void ApplyReadOnly()
    {
        this.Grid.ReadOnly = this.IsReadOnly;
        if (this.formulaBar is not null)
        {
            this.formulaBar.IsReadOnly = this.IsReadOnly;
        }

        this.IsReadOnlyChanged?.Invoke(this, EventArgs.Empty);
    }

    // Index of the active sheet in sheets.
    private int ActiveSheetIndex => this.sheets.IndexOf(this.vm);

    // ----- Public API (for host apps) -----

    public void NewSheet(int rows, int columns)
    {
        this.DetachActiveCell();
        this.vm = new SpreadsheetViewModel { Name = "Sheet1" };
        this.vm.Initialize(rows, columns);
        this.sheets.Clear();
        this.sheets.Add(this.vm);
        this.DataContext = this.vm;
        this.Grid.Sheet = this.vm;
        this.SetActiveCell(new CellAddress(0, 0));
        this.RebuildTabs();
        this.HookSheet();
    }

    public void LoadRows(IReadOnlyList<IReadOnlyList<string>> data)
    {
        this.DetachActiveCell();
        this.vm.LoadRows(data);
        this.Grid.Rebuild();
        this.SetActiveCell(new CellAddress(0, 0));
    }

    public IReadOnlyList<IReadOnlyList<string>> GetRows() => this.vm.ToRows();

    public void LoadCellData(IReadOnlyList<IReadOnlyList<CellData>> data)
    {
        this.DetachActiveCell();
        this.vm.LoadCellData(data);
        this.Grid.Rebuild();
        this.SetActiveCell(new CellAddress(0, 0));
    }

    public IReadOnlyList<IReadOnlyList<CellData>> GetCellData() => this.vm.ToCellData();

    // Loads a full single-sheet snapshot (resets the workbook to one sheet): cells, per-column
    // widths, merges, images, charts, data bars, and names.
    public void LoadSheet(SheetData sheet) =>
        this.LoadWorkbook(new[] { ("Sheet1", sheet) });

    // Replaces the whole workbook with the given named sheets (e.g. every worksheet of an
    // opened .xlsx) and opens the first one.
    public void LoadWorkbook(IReadOnlyList<(string Name, SheetData Data)> sheets)
    {
        if (sheets.Count == 0)
        {
            this.NewSheet(this.InitialRowCount, this.InitialColumnCount);
            return;
        }

        this.DetachActiveCell();
        this.sheets.Clear();
        foreach ((string name, SheetData data) in sheets)
        {
            SpreadsheetViewModel sheetVm = new() { Name = this.UniqueName(name) };
            ApplySheetData(sheetVm, data);
            this.sheets.Add(sheetVm);
        }

        this.ActivateSheet(0);
        this.RebuildTabs();
    }

    // Adds a new blank sheet at the end and switches to it.
    public void AddSheet()
    {
        SpreadsheetViewModel sheetVm = new() { Name = this.UniqueName("Sheet") };
        sheetVm.Initialize(this.InitialRowCount, this.InitialColumnCount);
        this.sheets.Add(sheetVm);
        this.ActivateSheet(this.sheets.Count - 1);
        this.RebuildTabs();
    }

    // Loads a SheetData snapshot into a view model (names set before cells so =Name resolves).
    private static void ApplySheetData(SpreadsheetViewModel vm, SheetData sheet)
    {
        vm.SetNames(sheet.Names);
        vm.LoadCellData(sheet.Cells);
        for (int i = 0; i < vm.Columns.Count && i < sheet.ColumnWidths.Count; i++)
        {
            if (sheet.ColumnWidths[i] > 0)
            {
                vm.Columns[i].Width = sheet.ColumnWidths[i];
            }
        }

        vm.SetMerges(sheet.Merges);
        vm.SetImages(sheet.Images);
        vm.SetDataBars(sheet.DataBars);
        vm.SetCharts(sheet.Charts);
    }

    // ----- Overflow menu: spreadsheet-specific actions + formulas section -----

    private MenuFlyout BuildOverflowMenu()
    {
        MenuFlyout menu = new();

        MenuItem chart = new() { Header = "Insert chart…" };
        chart.Click += async (_, _) => await this.ShowChartBuilderAsync();
        menu.Items.Add(chart);

        MenuItem print = new() { Header = "Print…" };
        print.Click += (_, _) => this.Print();
        menu.Items.Add(print);

        MenuItem autoSum = new() { Header = "AutoSum (Σ)" };
        autoSum.Click += (_, _) => this.AutoSum();
        menu.Items.Add(autoSum);

        MenuItem functions = new() { Header = "Functions" };
        foreach (MenuItem category in this.BuildFunctionMenuItems())
        {
            functions.Items.Add(category);
        }

        menu.Items.Add(functions);
        menu.Items.Add(new Separator());

        MenuItem addRow = new() { Header = "Add row" };
        addRow.Click += (_, _) => this.AddRow();
        menu.Items.Add(addRow);

        MenuItem addColumn = new() { Header = "Add column" };
        addColumn.Click += (_, _) => this.AddColumn();
        menu.Items.Add(addColumn);

        return menu;
    }

    // The function library as a list of category MenuItems (each with its functions), reused by the
    // overflow menu's "Functions" submenu.
    private List<MenuItem> BuildFunctionMenuItems()
    {
        List<MenuItem> categories = new();
        foreach ((string category, (string Name, string Signature)[] functions) in FunctionLibrary)
        {
            MenuItem group = new() { Header = category };
            foreach ((string name, string signature) in functions)
            {
                MenuItem item = new() { Header = name };
                ToolTip.SetTip(item, signature);
                string fn = name;
                item.Click += (_, _) => this.InsertFunction(fn);
                group.Items.Add(item);
            }

            categories.Add(group);
        }

        return categories;
    }

    private void AddRow()
    {
        if (this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint();
        this.vm.EnsureSize(this.vm.RowCount + 1, this.vm.ColumnCount);
        this.Grid.Rebuild();
    }

    private void AddColumn()
    {
        if (this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint();
        this.vm.EnsureSize(this.vm.RowCount, this.vm.ColumnCount + 1);
        this.Grid.Rebuild();
    }

    // ----- Printing -----

    // Renders the active sheet to a temporary HTML file and opens it in the default browser, where the
    // user prints (or saves as PDF) via the browser's own print dialog. Avalonia has no built-in print
    // API, so this HTML-preview route is the portable, dependency-free way to print. Best-effort.
    public void Print()
    {
        try
        {
            string directory = Path.Combine(Path.GetTempPath(), "SpreadsheetEditor");
            Directory.CreateDirectory(directory);
            string file = Path.Combine(directory, SanitizeFileName(this.vm.Name) + ".html");
            File.WriteAllText(file, this.ExportHtml());
            Process.Start(new ProcessStartInfo { FileName = file, UseShellExecute = true });
        }
        catch
        {
            // Never crash the editor if the temp file or default browser can't be opened.
        }
    }

    // Builds a printable HTML rendering of the active sheet's used range, with per-cell formatting
    // (bold/italic/underline/strikethrough, colors, alignment, font) and merged regions preserved.
    public string ExportHtml()
    {
        SpreadsheetViewModel sheet = this.vm;
        int lastRow = Math.Max(sheet.LastNonEmptyRow(), 0);
        int lastColumn = LastNonEmptyColumn(sheet, lastRow);

        // Merge maps: anchor cell -> its span, and the set of covered cells to skip.
        Dictionary<(int Row, int Column), MergeRange> anchors = new();
        HashSet<(int Row, int Column)> covered = new();
        foreach (MergeRange merge in sheet.Merges)
        {
            anchors[(merge.Row, merge.Column)] = merge;
            for (int r = merge.Row; r < merge.Row + merge.RowSpan; r++)
            {
                for (int c = merge.Column; c < merge.Column + merge.ColumnSpan; c++)
                {
                    if (r != merge.Row || c != merge.Column)
                    {
                        covered.Add((r, c));
                    }
                }
            }
        }

        StringBuilder html = new();
        string title = WebUtility.HtmlEncode(sheet.Name);
        html.Append("<!doctype html><html><head><meta charset=\"utf-8\"><title>").Append(title).Append("</title><style>");
        html.Append("body{font-family:Calibri,'Segoe UI',sans-serif;margin:16px;}");
        html.Append("h1{font-size:15pt;margin:0 0 10px;}");
        html.Append("table{border-collapse:collapse;table-layout:fixed;}");
        html.Append("td{border:1px solid #c8c8c8;padding:2px 6px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}");
        html.Append("</style></head><body><h1>").Append(title).Append("</h1><table><colgroup>");
        for (int c = 0; c <= lastColumn; c++)
        {
            double width = c < sheet.ColumnCount ? sheet.Columns[c].Width : sheet.DefaultColumnWidth;
            html.Append("<col style=\"width:").Append(((int)width).ToString(CultureInfo.InvariantCulture)).Append("px\">");
        }

        html.Append("</colgroup>");
        for (int r = 0; r <= lastRow; r++)
        {
            html.Append("<tr>");
            for (int c = 0; c <= lastColumn; c++)
            {
                if (covered.Contains((r, c)))
                {
                    continue;
                }

                CellViewModel? cell = sheet.GetCell(new CellAddress(r, c));
                html.Append("<td");
                if (anchors.TryGetValue((r, c), out MergeRange span))
                {
                    if (span.ColumnSpan > 1)
                    {
                        html.Append(" colspan=\"").Append(span.ColumnSpan).Append('"');
                    }

                    if (span.RowSpan > 1)
                    {
                        html.Append(" rowspan=\"").Append(span.RowSpan).Append('"');
                    }
                }

                string style = CellStyle(cell);
                if (style.Length > 0)
                {
                    html.Append(" style=\"").Append(style).Append('"');
                }

                html.Append('>').Append(WebUtility.HtmlEncode(cell?.DisplayText ?? string.Empty)).Append("</td>");
            }

            html.Append("</tr>");
        }

        html.Append("</table></body></html>");
        return html.ToString();
    }

    // The last column (0-based) holding any content within the used rows, so blank trailing columns
    // aren't printed.
    private static int LastNonEmptyColumn(SpreadsheetViewModel sheet, int lastRow)
    {
        int last = 0;
        for (int r = 0; r <= lastRow && r < sheet.RowCount; r++)
        {
            RowViewModel row = sheet.Rows[r];
            for (int c = row.Cells.Count - 1; c > last; c--)
            {
                if (!string.IsNullOrEmpty(row.Cells[c].DisplayText))
                {
                    last = c;
                    break;
                }
            }
        }

        return last;
    }

    // Inline CSS for a cell's whole-cell formatting.
    private static string CellStyle(CellViewModel? cell)
    {
        if (cell is null)
        {
            return string.Empty;
        }

        StringBuilder style = new();
        if (cell.Bold)
        {
            style.Append("font-weight:bold;");
        }

        if (cell.Italic)
        {
            style.Append("font-style:italic;");
        }

        if (cell.Underline || cell.Strikethrough)
        {
            style.Append("text-decoration:");
            if (cell.Underline)
            {
                style.Append("underline ");
            }

            if (cell.Strikethrough)
            {
                style.Append("line-through");
            }

            style.Append(';');
        }

        if (cell.ForegroundColor is { } fg)
        {
            style.Append("color:").Append(ToHex(fg)).Append(';');
        }

        if (cell.BackgroundColor is { } bg)
        {
            style.Append("background-color:").Append(ToHex(bg)).Append(';');
        }

        string align = cell.TextAlignment switch
        {
            TextAlignment.Center => "center",
            TextAlignment.Right => "right",
            _ => "left",
        };
        style.Append("text-align:").Append(align).Append(';');
        style.Append("font-size:").Append(((int)Math.Round(cell.FontSize * 3.0 / 4.0)).ToString(CultureInfo.InvariantCulture)).Append("pt;");
        if (!string.IsNullOrWhiteSpace(cell.FontFamily))
        {
            style.Append("font-family:'").Append(cell.FontFamily).Append("';");
        }

        return style.ToString();
    }

    private static string ToHex(Color color) => $"#{color.R:X2}{color.G:X2}{color.B:X2}";

    private static string SanitizeFileName(string name)
    {
        foreach (char invalid in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(invalid, '_');
        }

        return string.IsNullOrWhiteSpace(name) ? "sheet" : name;
    }

    // ----- Undo / redo (public, mirroring RichTextEditor.Undo/Redo) -----

    public void Undo()
    {
        if (!this.vm.CanUndo)
        {
            return;
        }

        CellAddress active = this.vm.ActiveCell;
        this.DetachActiveCell();
        this.vm.Undo();
        this.Grid.Rebuild();
        this.SetActiveCell(this.vm.IsValid(active) ? active : new CellAddress(0, 0));
        this.RefreshUndoButtons();
    }

    public void Redo()
    {
        if (!this.vm.CanRedo)
        {
            return;
        }

        CellAddress active = this.vm.ActiveCell;
        this.DetachActiveCell();
        this.vm.Redo();
        this.Grid.Rebuild();
        this.SetActiveCell(this.vm.IsValid(active) ? active : new CellAddress(0, 0));
        this.RefreshUndoButtons();
    }

    // Selects every cell in the active sheet (mirrors RichTextEditor.SelectAll).
    public void SelectAll() => this.Grid.SelectAll();

    // Subscribes to the active sheet's undo-state + content changes, unhooking the previous sheet so
    // events reflect only the active one.
    private void HookSheet()
    {
        if (this.hookedSheet is not null)
        {
            this.hookedSheet.UndoStateChanged -= this.RefreshUndoButtons;
            this.hookedSheet.DataChanged -= this.OnSheetContentChanged;
        }

        this.hookedSheet = this.vm;
        this.hookedSheet.UndoStateChanged += this.RefreshUndoButtons;
        this.hookedSheet.DataChanged += this.OnSheetContentChanged;
        this.RefreshUndoButtons();
    }

    private void OnSheetContentChanged() => this.ContentChanged?.Invoke(this, EventArgs.Empty);

    private void RefreshUndoButtons()
    {
        this.UndoButton.IsEnabled = this.vm.CanUndo;
        this.RedoButton.IsEnabled = this.vm.CanRedo;
    }

    // Inserts =NAME() into the active cell and puts the caret inside the parentheses
    // (in the formula bar) so the user can type the arguments.
    private void InsertFunction(string name)
    {
        if (this.activeCell is null || this.formulaBar is null || this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint();
        this.activeCell.Text = $"={name}()";
        this.Grid.SuspendEditor();
        this.formulaBar.Focus();
        this.formulaBar.CaretIndex = name.Length + 2; // just after "=NAME("
    }

    // Sums the contiguous numbers above the active cell (or to its left), into the active cell.
    private void AutoSum()
    {
        if (this.activeCell is null || this.formulaBar is null || this.IsReadOnly)
        {
            return;
        }

        // If a multi-cell range is selected, sum it; otherwise sum the run above/left of the cell.
        CellRange selection = this.Grid.SelectionRange;
        string? formula = selection.RowSpan > 1 || selection.ColumnSpan > 1
            ? $"=SUM({selection.ToA1()})"
            : this.SumRun(vertical: true) ?? this.SumRun(vertical: false);
        if (formula is null)
        {
            return;
        }

        this.vm.Checkpoint();
        this.activeCell.Text = formula;
        this.Grid.SuspendEditor();
        this.formulaBar.Focus();
        this.formulaBar.CaretIndex = formula.Length;
    }

    private string? SumRun(bool vertical)
    {
        CellAddress a = this.vm.ActiveCell;
        int i = (vertical ? a.Row : a.Column) - 1;
        while (i >= 0 && this.IsNumericCell(vertical ? new CellAddress(i, a.Column) : new CellAddress(a.Row, i)))
        {
            i--;
        }

        int start = i + 1;
        int end = (vertical ? a.Row : a.Column) - 1;
        if (start > end)
        {
            return null;
        }

        CellAddress from = vertical ? new CellAddress(start, a.Column) : new CellAddress(a.Row, start);
        CellAddress to = vertical ? new CellAddress(end, a.Column) : new CellAddress(a.Row, end);
        return $"=SUM({from.ToA1()}:{to.ToA1()})";
    }

    private bool IsNumericCell(CellAddress address)
    {
        CellViewModel? cell = this.vm.GetCell(address);
        return cell is not null && double.TryParse(cell.DisplayText,
            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _);
    }

    // Opens the chart builder and inserts the chart the user configures.
    private async Task ShowChartBuilderAsync()
    {
        if (TopLevel.GetTopLevel(this) is not Window owner)
        {
            return;
        }

        // Default to the user's selected range if it spans multiple cells, else the current region.
        CellRange selection = this.Grid.SelectionRange;
        CellRange defaultRange;
        if (selection.RowSpan > 1 || selection.ColumnSpan > 1)
        {
            defaultRange = selection;
        }
        else
        {
            MergeRange region = this.vm.GetCurrentRegion(this.vm.ActiveCell);
            defaultRange = new CellRange(region.Row, region.Column, region.RowSpan, region.ColumnSpan);
        }

        ChartBuilderResult? result = await ChartBuilderDialog.ShowAsync(owner, defaultRange.ToA1());
        if (result is null)
        {
            return;
        }

        if (CellRange.TryParse(result.Range, out CellRange range))
        {
            this.vm.InsertChart(result.Kind, range, result.FirstRowHeaders, result.FirstColCategories,
                result.Title, result.ShowLegend, result.ShowDataLabels,
                result.ValueAxisFormat, result.AxisMinimum, result.AxisMaximum, result.AxisMajorStep);
        }
    }

    // Opens the builder pre-filled with a chart's settings and applies the user's changes.
    private async Task EditChartAsync(SheetChart chart)
    {
        if (TopLevel.GetTopLevel(this) is not Window owner)
        {
            return;
        }

        CellRange range = chart.SourceRange ?? new CellRange(chart.AnchorRow, 0, 1, 1);
        ChartBuilderResult initial = new(chart.Kind, range.ToA1(), chart.Title,
            chart.FirstRowHeaders, chart.FirstColCategories, chart.ShowLegend, chart.ShowDataLabels,
            chart.ValueAxisFormat, chart.AxisMinimum, chart.AxisMaximum, chart.AxisMajorStep);
        ChartBuilderResult? result = await ChartBuilderDialog.ShowAsync(owner, range.ToA1(), initial);
        if (result is null)
        {
            return;
        }

        if (CellRange.TryParse(result.Range, out CellRange newRange))
        {
            this.vm.UpdateChart(chart, result.Kind, newRange, result.FirstRowHeaders, result.FirstColCategories,
                result.Title, result.ShowLegend, result.ShowDataLabels,
                result.ValueAxisFormat, result.AxisMinimum, result.AxisMaximum, result.AxisMajorStep);
        }
    }

    // The formula bar: a plain TextBox bound two-way to the active cell's text (its
    // DataContext is the active cell), so it shows and edits the cell content live.
    private void CreateFormulaBar()
    {
        this.formulaBar = new TextBox { VerticalContentAlignment = VerticalAlignment.Center };
        this.formulaBar.Bind(TextBox.TextProperty, new Binding(nameof(CellViewModel.Text))
        {
            Mode = BindingMode.TwoWay,
            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
        });

        // Clicking into the formula bar moves the single cursor here: hide the in-cell editor so the
        // caret only exists in one place, and record a pre-edit undo checkpoint.
        this.formulaBar.GotFocus += (_, _) =>
        {
            this.vm.Checkpoint();
            this.Grid.SuspendEditor();
        };
        this.EditorHost.Child = this.formulaBar;
    }

    // ----- Sheet tabs -----

    // Switches the editor to the sheet at index.
    private void ActivateSheet(int index)
    {
        if (index < 0 || index >= this.sheets.Count)
        {
            return;
        }

        this.DetachActiveCell();
        this.vm = this.sheets[index];
        this.DataContext = this.vm;
        this.Grid.Sheet = this.vm;            // rebinds + rebuilds the grid for this sheet
        this.SetActiveCell(new CellAddress(0, 0));
        this.HookSheet();
    }

    private void SwitchToSheet(SpreadsheetViewModel sheet)
    {
        if (ReferenceEquals(sheet, this.vm))
        {
            return;
        }

        this.ActivateSheet(this.sheets.IndexOf(sheet));
        this.RebuildTabs();
    }

    private void RenameSheet(SpreadsheetViewModel sheet, string name)
    {
        name = name.Trim();
        if (string.IsNullOrEmpty(name) || name.Equals(sheet.Name, StringComparison.Ordinal))
        {
            return;
        }

        sheet.Name = this.UniqueName(name, except: sheet);
        this.RebuildTabs();
    }

    private void DeleteSheet(SpreadsheetViewModel sheet)
    {
        if (this.sheets.Count <= 1)
        {
            return; // keep at least one sheet
        }

        int index = this.sheets.IndexOf(sheet);
        bool wasActive = ReferenceEquals(sheet, this.vm);
        this.sheets.Remove(sheet);
        if (wasActive)
        {
            this.ActivateSheet(Math.Min(index, this.sheets.Count - 1));
        }

        this.RebuildTabs();
    }

    // Returns desired, suffixed (e.g. "Sheet2") to stay unique among sheets.
    private string UniqueName(string desired, SpreadsheetViewModel? except = null)
    {
        desired = string.IsNullOrWhiteSpace(desired) ? "Sheet" : desired.Trim();
        bool Taken(string n) => this.sheets.Any(s => !ReferenceEquals(s, except) && s.Name.Equals(n, StringComparison.OrdinalIgnoreCase));
        if (!Taken(desired))
        {
            return desired;
        }

        for (int i = 2; ; i++)
        {
            string candidate = $"{desired}{i}";
            if (!Taken(candidate))
            {
                return candidate;
            }
        }
    }

    // Rebuilds the bottom tab strip from the sheet list (active tab highlighted).
    private void RebuildTabs()
    {
        this.SheetTabs.Children.Clear();
        foreach (SpreadsheetViewModel sheet in this.sheets)
        {
            this.SheetTabs.Children.Add(this.BuildTab(sheet));
        }
    }

    private Control BuildTab(SpreadsheetViewModel sheet)
    {
        bool active = ReferenceEquals(sheet, this.vm);
        bool dark = this.ActualThemeVariant == ThemeVariant.Dark;
        TextBlock text = new()
        {
            Text = sheet.Name,
            VerticalAlignment = VerticalAlignment.Center,
            FontWeight = active ? FontWeight.SemiBold : FontWeight.Normal,
        };
        // Active tab reads as the "front" sheet: an elevated surface (light or dark) with a blue top
        // edge; its label uses the inherited theme foreground so it stays legible in both themes.
        IBrush activeBackground = dark ? new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C)) : Brushes.White;
        Button tab = new()
        {
            Content = text,
            Focusable = false,
            Padding = new Thickness(10, 3),
            Background = active ? activeBackground : Brushes.Transparent,
            BorderThickness = new Thickness(1, active ? 2 : 1, 1, 0),
            BorderBrush = active ? new SolidColorBrush(Color.FromRgb(0x1A, 0x73, 0xE8)) : Brushes.Transparent,
        };
        tab.Click += (_, _) => this.SwitchToSheet(sheet);

        // Double-click a tab to rename it inline.
        tab.AddHandler(PointerPressedEvent, (_, e) =>
        {
            if (e.ClickCount >= 2)
            {
                this.BeginRenameTab(sheet, tab);
            }
        }, Avalonia.Interactivity.RoutingStrategies.Tunnel);

        // Right-click: Rename / Delete.
        MenuFlyout menu = new();
        MenuItem rename = new() { Header = "Rename" };
        rename.Click += (_, _) => this.BeginRenameTab(sheet, tab);
        MenuItem delete = new() { Header = "Delete", IsEnabled = this.sheets.Count > 1 };
        delete.Click += (_, _) => this.DeleteSheet(sheet);
        menu.Items.Add(rename);
        menu.Items.Add(delete);
        tab.ContextFlyout = menu;
        return tab;
    }

    // Swaps a tab's label for an inline editor to rename the sheet.
    private void BeginRenameTab(SpreadsheetViewModel sheet, Button tab)
    {
        TextBox box = new()
        {
            Text = sheet.Name,
            MinWidth = 70,
            Padding = new Thickness(2, 0),
        };
        tab.Content = box;
        void Commit()
        {
            this.RenameSheet(sheet, box.Text ?? sheet.Name);
            this.RebuildTabs();
        }

        box.KeyDown += (_, e) =>
        {
            if (e.Key == Avalonia.Input.Key.Enter) { Commit(); e.Handled = true; }
            else if (e.Key == Avalonia.Input.Key.Escape) { this.RebuildTabs(); e.Handled = true; }
        };
        box.LostFocus += (_, _) => Commit();
        Dispatcher.UIThread.Post(() => { box.Focus(); box.SelectAll(); });
    }

    // ----- Active cell -----

    private void SetActiveCell(CellAddress address)
    {
        if (!this.vm.IsValid(address))
        {
            return;
        }

        if (this.activeCell is not null && this.vm.ActiveCell.Equals(address))
        {
            return;
        }

        this.DetachActiveCell();

        this.vm.ActiveCell = address;
        this.activeCell = this.vm.GetCell(address);
        this.NameBox.Text = address.ToA1();

        if (this.activeCell is null)
        {
            return;
        }

        this.activeCell.IsActive = true;
        if (this.formulaBar is not null)
        {
            this.formulaBar.DataContext = this.activeCell;
        }

        this.SyncToolbar(this.activeCell);
        this.SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    private void DetachActiveCell()
    {
        if (this.formulaBar is not null)
        {
            this.formulaBar.DataContext = null;
        }

        if (this.activeCell is null)
        {
            return;
        }

        this.activeCell.IsActive = false;
        this.activeCell = null;
    }

    // ----- Toolbar formatting (applies to every cell in the selection) -----

    private void ApplyToSelection(Action<CellViewModel> apply)
    {
        if (this.suppressToolbar || this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint(); // whole-selection formatting is one undo step
        foreach (CellViewModel cell in this.Grid.SelectionCells())
        {
            apply(cell);
        }
    }

    private void ApplyForeground(Color color)
    {
        this.ForegroundSwatch.Background = new SolidColorBrush(color);
        if (this.suppressToolbar)
        {
            return;
        }

        // Route through the editor: per-character when editing, else whole-selection.
        this.Grid.ApplyForegroundToCell(color);
    }

    private void ApplyBackground(Color color)
    {
        this.BackgroundSwatch.Background = new SolidColorBrush(color);
        this.ApplyToSelection(c => c.BackgroundColor = color);
    }

    private void SyncToolbar(CellViewModel cell)
    {
        this.suppressToolbar = true;
        try
        {
            this.BoldButton.IsChecked = cell.Bold;
            this.ItalicButton.IsChecked = cell.Italic;
            this.UnderlineButton.IsChecked = cell.Underline;
            this.StrikethroughButton.IsChecked = cell.Strikethrough;
            this.ForegroundSwatch.Background = cell.Foreground ?? Brushes.Black;
            this.BackgroundSwatch.Background = cell.Background ?? Brushes.Transparent;

            this.FontFamilyBox.SelectedItem = cell.FontFamily ?? DefaultFontName;
            this.FontSizeBox.SelectedItem = (int)Math.Round(cell.FontSize * 3.0 / 4.0); // px → pt
            this.AlignLeftButton.IsChecked = cell.TextAlignment == TextAlignment.Left;
            this.AlignCenterButton.IsChecked = cell.TextAlignment == TextAlignment.Center;
            this.AlignRightButton.IsChecked = cell.TextAlignment == TextAlignment.Right;
        }
        finally
        {
            this.suppressToolbar = false;
        }
    }

    private void SetupFontControls()
    {
        this.FontFamilyBox.ItemsSource = FontNames;
        this.FontSizeBox.ItemsSource = FontSizes;
        this.FontFamilyBox.SelectionChanged += (_, _) =>
        {
            if (!this.suppressToolbar && this.FontFamilyBox.SelectedItem is string name)
            {
                this.ApplyToSelection(c => c.FontFamily = name);
            }
        };
        this.FontSizeBox.SelectionChanged += (_, _) =>
        {
            if (!this.suppressToolbar && this.FontSizeBox.SelectedItem is int pt)
            {
                this.ApplyToSelection(c => c.FontSize = pt * 4.0 / 3.0); // pt → px
            }
        };
    }

    // Reflects the current B/I/U state (editor selection while editing, else whole-cell) on the toolbar.
    private void SyncFormatButtons()
    {
        (bool bold, bool italic, bool underline, bool strikethrough) = this.Grid.FormatState();
        this.suppressToolbar = true;
        try
        {
            this.BoldButton.IsChecked = bold;
            this.ItalicButton.IsChecked = italic;
            this.UnderlineButton.IsChecked = underline;
            this.StrikethroughButton.IsChecked = strikethrough;
        }
        finally
        {
            this.suppressToolbar = false;
        }
    }

    private void SetAlignment(TextAlignment alignment)
    {
        this.ApplyToSelection(c => c.TextAlignment = alignment);
        this.suppressToolbar = true;
        this.AlignLeftButton.IsChecked = alignment == TextAlignment.Left;
        this.AlignCenterButton.IsChecked = alignment == TextAlignment.Center;
        this.AlignRightButton.IsChecked = alignment == TextAlignment.Right;
        this.suppressToolbar = false;
    }

    private void BuildPalette(WrapPanel panel, Action<Color> apply)
    {
        foreach (Color color in Palette)
        {
            Button swatch = new()
            {
                Width = 28,
                Height = 28,
                Margin = new Thickness(1),
                Padding = new Thickness(0),
                Background = new SolidColorBrush(color),
                BorderThickness = new Thickness(1),
                BorderBrush = Brushes.Gray,
                Focusable = false,
            };
            Color captured = color;
            swatch.Click += (_, _) => apply(captured);
            panel.Children.Add(swatch);
        }
    }

    // ----- Insert special characters (currency + math symbols) -----

    private static readonly string[] CurrencySymbols =
    {
        "$", "€", "£", "¥", "¢", "₹", "₩",
        "₽", "₪", "₺", "₫", "฿", "₦", "₱",
    };

    private static readonly string[] MathSymbols =
    {
        "+", "−", "×", "÷", "=", "≠", "≈", "±",
        "√", "∑", "∏", "∫", "∞", "≤", "≥",
        "π", "µ", "°", "‰", "²", "³",
        "½", "¼", "¾", "·", "Δ",
    };

    private Flyout BuildSymbolFlyout()
    {
        StackPanel root = new() { Spacing = 6, MaxWidth = 270 };
        root.Children.Add(this.SymbolSection("Currency", CurrencySymbols));
        root.Children.Add(this.SymbolSection("Math", MathSymbols));
        return new Flyout { Content = root };
    }

    private Control SymbolSection(string title, string[] symbols)
    {
        StackPanel section = new() { Spacing = 2 };
        section.Children.Add(new TextBlock { Text = title, FontWeight = FontWeight.SemiBold, Margin = new Thickness(2, 0) });
        WrapPanel grid = new() { MaxWidth = 264 };
        foreach (string symbol in symbols)
        {
            Button button = new()
            {
                Content = symbol,
                Width = 30,
                Height = 28,
                Margin = new Thickness(1),
                Padding = new Thickness(0),
                Focusable = false,
                FontSize = 15,
            };
            ToolTip.SetTip(button, symbol);
            string captured = symbol;
            button.Click += (_, _) => this.InsertSymbol(captured);
            grid.Children.Add(button);
        }

        section.Children.Add(grid);
        return section;
    }

    // Appends a symbol to the active cell's content (via the formula bar), as one undo step.
    private void InsertSymbol(string symbol)
    {
        if (this.activeCell is null || this.formulaBar is null || this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint();
        this.Grid.SuspendEditor();
        this.formulaBar.Focus();
        string text = this.formulaBar.Text ?? string.Empty;
        this.formulaBar.Text = text + symbol;
        this.formulaBar.CaretIndex = (this.formulaBar.Text ?? string.Empty).Length;
    }
}
