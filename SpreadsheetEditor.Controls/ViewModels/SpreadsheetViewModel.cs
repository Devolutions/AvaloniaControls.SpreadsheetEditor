namespace SpreadsheetEditor.Controls.ViewModels;

using global::SpreadsheetEditor.Controls.Localization;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using Avalonia.Collections;
using Avalonia.Media;

using CommunityToolkit.Mvvm.ComponentModel;

using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Models;

// Source of truth for the sheet: the row/column grid plus the active-cell address. Kept free of
// any view (DataGrid / RichTextEditor) types so the rendering layer can change independently.
public partial class SpreadsheetViewModel : ObservableObject
{
    public AvaloniaList<ColumnViewModel> Columns { get; } = new();

    public AvaloniaList<RowViewModel> Rows { get; } = new();

    // The sheet's tab name (shown in the bottom tab strip).
    [ObservableProperty]
    private string name = EditorText.Get(EditorString.Sheet) + "1";

    // Optional color shown on the worksheet tab (imported from Excel and preserved on save).
    [ObservableProperty]
    private Color? tabColor;

    [ObservableProperty]
    private CellAddress activeCell;

    public int RowCount => this.Rows.Count;

    public int ColumnCount => this.Columns.Count;

    public double DefaultColumnWidth { get; set; } = 90;

    // Merged cell regions (from an imported sheet). Rendered by the grid that supports spanning.
    public IReadOnlyList<MergeRange> Merges { get; private set; } = new List<MergeRange>();

    public void SetMerges(IReadOnlyList<MergeRange> merges) => this.Merges = merges;

    // Pictures embedded in the imported sheet. Rendered as overlays by the grid.
    public IReadOnlyList<SheetImage> Images { get; private set; } = new List<SheetImage>();

    public void SetImages(IReadOnlyList<SheetImage> images) => this.Images = images;

    // Conditional-formatting data bars (proportional in-cell bars).
    public IReadOnlyList<SheetDataBar> DataBars { get; private set; } = new List<SheetDataBar>();

    public void SetDataBars(IReadOnlyList<SheetDataBar> dataBars) => this.DataBars = dataBars;

    public IReadOnlyList<SheetCellValueRule> CellValueRules { get; private set; } = new List<SheetCellValueRule>();

    public void SetCellValueRules(IReadOnlyList<SheetCellValueRule> rules) => this.CellValueRules = rules;

    public CellRange? AutoFilterRange { get; private set; }

    public void SetAutoFilterRange(CellRange? range) => this.AutoFilterRange = range;

    private readonly HashSet<RowViewModel> sourceHiddenRows = new();

    public void SetHiddenRows(IReadOnlyList<int> hiddenRows)
    {
        this.sourceHiddenRows.Clear();
        foreach (RowViewModel row in this.Rows)
        {
            row.IsHidden = false;
        }
        foreach (int row in hiddenRows)
        {
            if (row >= 0 && row < this.RowCount)
            {
                this.Rows[row].IsHidden = true;
                this.sourceHiddenRows.Add(this.Rows[row]);
            }
        }
    }

    public void ShowAllRows()
    {
        this.columnFilters.Clear();
        this.sourceHiddenRows.Clear();
        foreach (RowViewModel row in this.Rows)
        {
            row.IsHidden = false;
        }
        this.ViewChanged?.Invoke();
    }

    public event Action? ViewChanged;

    // Formatting, comments and links change the saved workbook without changing formula inputs.
    // Notify the host (autosave/dirty tracking) without recalculating every formula.
    public void NotifyMetadataChanged() => this.ViewChanged?.Invoke();

    // Charts embedded in the sheet. The grid renders these as live overlays.
    public AvaloniaList<SheetChart> Charts { get; } = new();

    internal event Action? ChartLayoutChanged;

    // Raised after a recalculation so chart overlays can refresh their data.
    public event Action? DataChanged;

    private readonly FormulaEngine engine;
    private readonly HashSet<CellAddress> formulaCells = new();

    private bool suspendRecalc;

    private Dictionary<string, string> names = new(StringComparer.OrdinalIgnoreCase);
    private Func<string, CellAddress, CellValue?>? workbookCellResolver;
    private Func<IReadOnlyList<CellAddress>, bool, bool>? editRecalculationGate;

    private readonly Dictionary<string, HashSet<CellRange>> formulaInputRanges = new(StringComparer.OrdinalIgnoreCase);
    private long formulaVersion;
    private long indexedFormulaVersion = -1;
    private bool uncertainFormulaInputs;

    public SpreadsheetViewModel() => this.engine = new FormulaEngine(this.ReadRaw, this.ResolveName, this.ResolveCellReference);

    // Defined names mapped to their formula text (e.g. NetIncome -> "O36",
    // TotalMonthlyIncome -> "SUM('Sheet'!$C$7:$C$10)") so formulas like =NetIncome resolve.
    public void SetNames(IReadOnlyDictionary<string, string> names)
    {
        this.names = new Dictionary<string, string>(names, StringComparer.OrdinalIgnoreCase);
        this.CalculationVersion++;
        this.formulaVersion++;
    }

    // The defined names (name -> RefersTo formula/range), so a save can persist them back to the
    // .xlsx — otherwise formulas that reference a name become #NAME? on the next open.
    public IReadOnlyDictionary<string, string> Names => this.names;

    private string? ResolveName(string name) => this.names.TryGetValue(name, out string? f) ? f : null;

    internal void SetWorkbookCellResolver(Func<string, CellAddress, CellValue?>? resolver) =>
        this.workbookCellResolver = resolver;

    // Lets the workbook host postpone a recalculation until lazily loaded cross-sheet references are
    // available. Standalone view models keep their normal immediate-recalculation behavior.
    internal void SetEditRecalculationGate(Func<IReadOnlyList<CellAddress>, bool, bool>? gate) => this.editRecalculationGate = gate;

    internal CellValue EvaluateCell(CellAddress address) => this.engine.Evaluate(address);

    internal long CalculationVersion { get; private set; }

    internal CellAddress[] FormulaAddresses() => this.formulaCells.ToArray();

    internal void ApplyCalculatedValue(CellAddress address, CellValue value, string display)
    {
        CellViewModel? cell = this.GetCell(address);
        if (cell is null)
        {
            return;
        }
        cell.CachedNumericValue = value.Kind == CellValueKind.Number ? value.Number : null;
        cell.CachedFormulaValue = value;
        cell.ComputedValue = display;
    }

    internal void CompleteBackgroundCalculation()
    {
        this.engine.InvalidateCache();
        this.DataChanged?.Invoke();
    }

    internal IReadOnlySet<string> ReferencedSheetNames()
    {
        this.IndexFormulaInputs();
        return new HashSet<string>(this.formulaInputRanges.Keys.Where(name =>
            name.Length > 0 && name != "*" && !name.Equals(this.Name, StringComparison.OrdinalIgnoreCase)), StringComparer.OrdinalIgnoreCase);
    }

    internal bool HasVolatileFormulas => this.formulaInputRanges.ContainsKey("*");

    internal bool ReferencesInput(string sheetName, CellAddress address)
    {
        if (this.indexedFormulaVersion != this.formulaVersion)
        {
            // Import builds this index before publication. A changed formula invalidates it;
            // be conservative until the next index build, rather than parsing a large sheet on a keystroke.
            return this.formulaCells.Count > 0;
        }
        bool Contains(string key) => this.formulaInputRanges.TryGetValue(key, out HashSet<CellRange>? ranges) &&
            ranges.Any(range => address.Row >= range.Row && address.Row < range.Row + range.RowSpan &&
                address.Column >= range.Column && address.Column < range.Column + range.ColumnSpan);
        return this.uncertainFormulaInputs || Contains("*") || Contains(sheetName) ||
            (sheetName.Equals(this.Name, StringComparison.OrdinalIgnoreCase) && Contains(string.Empty));
    }

    private void IndexFormulaInputs()
    {
        if (this.indexedFormulaVersion == this.formulaVersion)
        {
            return;
        }
        this.formulaInputRanges.Clear();
        this.uncertainFormulaInputs = false;
        foreach (CellAddress address in this.formulaCells)
        {
            try
            {
                string text = this.GetCell(address)?.Text ?? string.Empty;
                if (!FormulaEngine.IsFormula(text))
                {
                    continue;
                }
                foreach ((string? sheet, CellRange range) in FormulaParser.ReferencedRanges(text[1..], this.ResolveName))
                {
                    string key = sheet ?? string.Empty;
                    if (!this.formulaInputRanges.TryGetValue(key, out HashSet<CellRange>? ranges))
                    {
                        this.formulaInputRanges[key] = ranges = new HashSet<CellRange>();
                    }
                    ranges.Add(range);
                }
            }
            catch (FormulaParseException)
            {
                // Unknown inputs must be treated conservatively rather than skipping a needed update.
                this.uncertainFormulaInputs = true;
            }
        }
        this.indexedFormulaVersion = this.formulaVersion;
    }

    public SpreadsheetViewModel(int rows, int columns) : this() => this.Initialize(rows, columns);

    private string ReadRaw(CellAddress a) => this.IsValid(a) ? this.Rows[a.Row].Cells[a.Column].Text : string.Empty;

    private CellValue ResolveCellReference(string? sheetName, CellAddress address)
    {
        if (string.IsNullOrWhiteSpace(sheetName) || sheetName.Equals(this.Name, StringComparison.OrdinalIgnoreCase))
        {
            return this.engine.Evaluate(address);
        }

        return this.workbookCellResolver?.Invoke(sheetName, address) ?? CellValue.FromError(FormulaErrors.Ref);
    }

    // Runs a bulk edit with recalculation suspended, then recalculates once at the end.
    public void RunBatch(Action action)
    {
        bool wasSuspended = this.suspendRecalc;
        this.suspendRecalc = true;
        try
        {
            action();
        }
        finally
        {
            this.suspendRecalc = wasSuspended;
        }

        if (!wasSuspended) this.RequestRecalculation();
    }

    private void RequestRecalculation()
    {
        // Bulk edits/undo must use the same asynchronous workbook path as single-cell edits.
        // Formula changes are conservative here because they can introduce new dependencies.
        CellAddress[] changed = this.batchedChangedCells.ToArray();
        this.batchedChangedCells.Clear();
        if (this.editRecalculationGate?.Invoke(changed, true) ?? true)
        {
            this.Recalculate();
        }
    }

    // Recomputes only formula cells. Literal cells are resolved on demand by formulas, so walking the
    // entire grid here would make typing cost proportional to worksheet size.
    public void Recalculate()
    {
        this.engine.Recalculate(this.formulaCells, (addr, value) =>
        {
            CellViewModel cell = this.Rows[addr.Row].Cells[addr.Column];
            cell.CachedNumericValue = value is { Kind: CellValueKind.Number } numeric ? numeric.Number : null;
            cell.CachedFormulaValue = value;
            cell.ComputedValue = value is { } v ? FormatComputed(v, cell.NumberFormat) : null;
        });

        this.DataChanged?.Invoke();
    }

    // Excel-style "current region": the contiguous block of non-empty cells containing
    // origin, grown outward while a bordering row/column has any content.
    public MergeRange GetCurrentRegion(CellAddress origin)
    {
        if (!this.IsValid(origin))
        {
            return new MergeRange(0, 0, 1, 1);
        }

        int r1 = origin.Row, r2 = origin.Row, c1 = origin.Column, c2 = origin.Column;
        bool Filled(int r, int c) => !string.IsNullOrWhiteSpace(this.Rows[r].Cells[c].DisplayText);
        bool RowHasContent(int r, int from, int to) => Enumerable.Range(from, to - from + 1).Any(c => Filled(r, c));
        bool ColHasContent(int c, int from, int to) => Enumerable.Range(from, to - from + 1).Any(r => Filled(r, c));

        bool grew;
        do
        {
            grew = false;
            if (r1 > 0 && RowHasContent(r1 - 1, c1, c2))
            {
                r1--;
                grew = true;
            }

            if (r2 < this.RowCount - 1 && RowHasContent(r2 + 1, c1, c2))
            {
                r2++;
                grew = true;
            }

            if (c1 > 0 && ColHasContent(c1 - 1, r1, r2))
            {
                c1--;
                grew = true;
            }

            if (c2 < this.ColumnCount - 1 && ColHasContent(c2 + 1, r1, r2))
            {
                c2++;
                grew = true;
            }
        }
        while (grew);

        return new MergeRange(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }

    // Replaces the sheet's charts (e.g. from an imported file).
    public void SetCharts(IReadOnlyList<SheetChart> charts)
    {
        this.Charts.Clear();
        this.Charts.AddRange(charts);
    }

    // Quick insert: chart the current region around the active cell with sensible defaults.
    public void InsertChart(ChartKind kind)
    {
        MergeRange region = this.GetCurrentRegion(this.ActiveCell);
        bool multiColumn = region.ColumnSpan > 1;
        CellViewModel? firstValue = this.GetCell(new CellAddress(region.Row, multiColumn ? region.Column + 1 : region.Column));
        bool header = region.RowSpan > 1 && firstValue is not null && !LooksNumeric(firstValue.DisplayText);
        this.InsertChart(kind, new CellRange(region.Row, region.Column, region.RowSpan, region.ColumnSpan),
            firstRowHeaders: header, firstColCategories: multiColumn, title: null, showLegend: true, showDataLabels: false);
    }

    // Builds a chart from a user-chosen range and options (the "chart builder"). The first
    // column (when firstColCategories) is category labels; each remaining column is
    // a series; a header row supplies series titles. Placed to the right of the data.
    public void InsertChart(ChartKind kind, CellRange range, bool firstRowHeaders, bool firstColCategories,
        string? title, bool showLegend, bool showDataLabels,
        AxisNumberFormat valueAxisFormat = AxisNumberFormat.Number,
        double? axisMin = null, double? axisMax = null, double? axisStep = null)
    {
        (List<ChartSeriesSpec> series, CellRange? categories) = this.BuildChartData(range, firstRowHeaders, firstColCategories);
        this.Charts.Add(new SheetChart
        {
            Kind = kind,
            Series = series,
            Categories = categories,
            Title = string.IsNullOrWhiteSpace(title) ? null : title,
            ShowLegend = showLegend,
            ShowDataLabels = showDataLabels,
            ValueAxisFormat = valueAxisFormat,
            AxisMinimum = axisMin,
            AxisMaximum = axisMax,
            AxisMajorStep = axisStep,
            SourceRange = range,
            FirstRowHeaders = firstRowHeaders,
            FirstColCategories = firstColCategories,
            AnchorRow = range.Row,
            AnchorColumn = Math.Min(range.Column + range.ColumnSpan, Math.Max(0, this.ColumnCount - 1)),
            OffsetX = 8,
        });
    }

    // Removes a chart from the sheet (the overlay re-lays-out via the collection change).
    public void RemoveChart(SheetChart chart) => this.Charts.Remove(chart);

    // Moves a chart so its top-left corner snaps into the given cell (keeping its current pixel
    // size), then notifies listeners so the overlay re-lays-out.
    public void MoveChart(SheetChart chart, int anchorRow, int anchorColumn, double width, double height)
    {
        chart.MoveTo(anchorRow, anchorColumn, width, height);
        this.ChartLayoutChanged?.Invoke();
        this.NotifyMetadataChanged();
    }

    // Resizes a chart to an explicit pixel size (keeping its anchor) and re-lays-out the overlay.
    public void ResizeChart(SheetChart chart, double width, double height)
    {
        chart.ResizeTo(width, height);
        this.ChartLayoutChanged?.Invoke();
        this.NotifyMetadataChanged();
    }

    // Replaces an existing chart with new settings, keeping its on-sheet placement/size.
    // Re-applies only presentation settings (kind/title/legend/labels/axis), keeping the chart's
    // existing series + categories. Used when editing a chart whose series are explicit ranges (e.g.
    // imported charts with no single SourceRange), so editing settings never wipes its data.
    public void UpdateChartSettings(SheetChart existing, ChartKind kind, string? title, bool showLegend,
        bool showDataLabels, AxisNumberFormat valueAxisFormat = AxisNumberFormat.Number,
        double? axisMin = null, double? axisMax = null, double? axisStep = null)
    {
        int index = this.Charts.IndexOf(existing);
        if (index < 0)
        {
            return;
        }

        this.Charts[index] = new SheetChart
        {
            Kind = kind,
            Series = existing.Series,
            Categories = existing.Categories,
            CachedCategories = existing.CachedCategories,
            Title = string.IsNullOrWhiteSpace(title) ? null : title,
            ShowLegend = showLegend,
            ShowDataLabels = showDataLabels,
            ValueAxisFormat = valueAxisFormat,
            AxisMinimum = axisMin,
            AxisMaximum = axisMax,
            AxisMajorStep = axisStep,
            SourceRange = existing.SourceRange,
            FirstRowHeaders = existing.FirstRowHeaders,
            FirstColCategories = existing.FirstColCategories,
            AnchorRow = existing.AnchorRow,
            AnchorColumn = existing.AnchorColumn,
            OffsetX = existing.OffsetX,
            OffsetY = existing.OffsetY,
            ToRow = existing.ToRow,
            ToColumn = existing.ToColumn,
            ToOffsetX = existing.ToOffsetX,
            ToOffsetY = existing.ToOffsetY,
            Width = existing.Width,
            Height = existing.Height,
        };
    }

    // Retargets an imported chart's explicit source ranges: each series keeps its title but takes a new
    // value range (blank/invalid entries keep the current one), and the category range is updated. Cached
    // values are cleared so the chart reads the new ranges live. Presentation settings are applied too.
    public void UpdateChartSource(SheetChart existing, IReadOnlyList<string>? seriesRanges, string? categoryRange,
        ChartKind kind, string? title, bool showLegend, bool showDataLabels,
        AxisNumberFormat valueAxisFormat = AxisNumberFormat.Number,
        double? axisMin = null, double? axisMax = null, double? axisStep = null)
    {
        int index = this.Charts.IndexOf(existing);
        if (index < 0)
        {
            return;
        }

        List<ChartSeriesSpec> series = new();
        for (int i = 0; i < existing.Series.Count; i++)
        {
            ChartSeriesSpec spec = existing.Series[i];
            IReadOnlyList<CellRange> parsed = seriesRanges is not null && i < seriesRanges.Count
                ? ChartSeriesSpec.ParseRanges(seriesRanges[i])
                : new List<CellRange>();
            if (parsed.Count > 0)
            {
                series.Add(new ChartSeriesSpec(spec.Title, parsed)); // retargeted (may be several cells) → read live
            }
            else
            {
                series.Add(new ChartSeriesSpec(spec.Title, spec.Ranges, spec.CachedValues)); // unchanged → keep
            }
        }

        // Category field: blank keeps the existing categories/labels (auto); a valid range reads live;
        // anything else is treated as comma/semicolon-separated literal labels the user typed.
        CellRange? categories;
        IReadOnlyList<string>? cachedCategories;
        if (string.IsNullOrWhiteSpace(categoryRange))
        {
            categories = existing.Categories;
            cachedCategories = existing.CachedCategories;
        }
        else if (CellRange.TryParse(categoryRange, out CellRange parsedCategories))
        {
            categories = parsedCategories;
            cachedCategories = null;
        }
        else
        {
            categories = null;
            cachedCategories = categoryRange
                .Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(static s => s.Trim())
                .Where(static s => s.Length > 0)
                .ToList();
        }

        this.Charts[index] = new SheetChart
        {
            Kind = kind,
            Series = series,
            Categories = categories,
            CachedCategories = cachedCategories,
            Title = string.IsNullOrWhiteSpace(title) ? null : title,
            ShowLegend = showLegend,
            ShowDataLabels = showDataLabels,
            ValueAxisFormat = valueAxisFormat,
            AxisMinimum = axisMin,
            AxisMaximum = axisMax,
            AxisMajorStep = axisStep,
            SourceRange = existing.SourceRange,
            FirstRowHeaders = existing.FirstRowHeaders,
            FirstColCategories = existing.FirstColCategories,
            AnchorRow = existing.AnchorRow,
            AnchorColumn = existing.AnchorColumn,
            OffsetX = existing.OffsetX,
            OffsetY = existing.OffsetY,
            ToRow = existing.ToRow,
            ToColumn = existing.ToColumn,
            ToOffsetX = existing.ToOffsetX,
            ToOffsetY = existing.ToOffsetY,
            Width = existing.Width,
            Height = existing.Height,
        };
    }

    public void UpdateChart(SheetChart existing, ChartKind kind, CellRange range, bool firstRowHeaders,
        bool firstColCategories, string? title, bool showLegend, bool showDataLabels,
        AxisNumberFormat valueAxisFormat = AxisNumberFormat.Number,
        double? axisMin = null, double? axisMax = null, double? axisStep = null)
    {
        int index = this.Charts.IndexOf(existing);
        if (index < 0)
        {
            return;
        }

        (List<ChartSeriesSpec> series, CellRange? categories) = this.BuildChartData(range, firstRowHeaders, firstColCategories);
        this.Charts[index] = new SheetChart
        {
            Kind = kind,
            Series = series,
            Categories = categories,
            Title = string.IsNullOrWhiteSpace(title) ? null : title,
            ShowLegend = showLegend,
            ShowDataLabels = showDataLabels,
            ValueAxisFormat = valueAxisFormat,
            AxisMinimum = axisMin,
            AxisMaximum = axisMax,
            AxisMajorStep = axisStep,
            SourceRange = range,
            FirstRowHeaders = firstRowHeaders,
            FirstColCategories = firstColCategories,
            AnchorRow = existing.AnchorRow,
            AnchorColumn = existing.AnchorColumn,
            OffsetX = existing.OffsetX,
            OffsetY = existing.OffsetY,
            ToRow = existing.ToRow,
            ToColumn = existing.ToColumn,
            ToOffsetX = existing.ToOffsetX,
            ToOffsetY = existing.ToOffsetY,
            Width = existing.Width,
            Height = existing.Height,
        };
    }

    // Builds series + category specs from a range and the header/category options.
    private (List<ChartSeriesSpec> Series, CellRange? Categories) BuildChartData(CellRange range, bool firstRowHeaders, bool firstColCategories)
    {
        bool multiColumn = range.ColumnSpan > 1 && firstColCategories;
        int valueColStart = multiColumn ? range.Column + 1 : range.Column;
        bool header = firstRowHeaders && range.RowSpan > 1;
        int dataRow = header ? range.Row + 1 : range.Row;
        int dataRowSpan = Math.Max(1, range.RowSpan - (header ? 1 : 0));

        List<ChartSeriesSpec> series = new();
        for (int c = valueColStart; c < range.Column + range.ColumnSpan; c++)
        {
            string seriesTitle = header ? this.GetCell(new CellAddress(range.Row, c))?.DisplayText ?? "" : EditorText.Format(EditorString.SeriesNumber, c - valueColStart + 1);
            series.Add(new ChartSeriesSpec(seriesTitle, new CellRange(dataRow, c, dataRowSpan, 1)));
        }

        CellRange? categories = multiColumn ? new CellRange(dataRow, range.Column, dataRowSpan, 1) : null;
        return (series, categories);
    }

    private static bool LooksNumeric(string s) =>
        double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _);

    // Renders a computed value: numbers use the cell's Excel number format; text/bool/error
    // use their plain representation.
    private static string FormatComputed(CellValue value, string? numberFormat) =>
        value.Kind == CellValueKind.Number
            ? NumberFormatter.Format(value.Number, numberFormat)
            : value.Display();

    // Subscribes a cell so editing its text triggers a recalculation (unless suspended for a bulk load).
    private void HookCell(CellViewModel cell, RowViewModel row, int column)
    {
        cell.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName != nameof(CellViewModel.Text))
            {
                return;
            }

            this.CalculationVersion++;
            this.engine.InvalidateCache();

            // Row objects move during insert/delete. Resolve their current position instead of
            // retaining the address they had when the cell was first created.
            CellAddress address = new(row.Number - 1, column);
            bool formulaChanged = this.formulaCells.Contains(address) || FormulaEngine.IsFormula(cell.Text);
            if (formulaChanged)
            {
                this.formulaVersion++;
            }

            if (FormulaEngine.IsFormula(cell.Text))
            {
                this.formulaCells.Add(address);
            }
            else
            {
                this.formulaCells.Remove(address);
            }

            if (this.suspendRecalc && this.editRecalculationGate is not null)
            {
                this.batchedChangedCells.Add(address);
            }
            if (!this.suspendRecalc)
            {
                if (this.editRecalculationGate?.Invoke(new[] { address }, formulaChanged) ?? true)
                {
                    this.Recalculate();
                }
            }
        };
    }

    private readonly HashSet<CellAddress> batchedChangedCells = new();

    // (Re)builds an empty sheet of the given size.
    public void Initialize(int rows, int columns)
    {
        this.CalculationVersion++;
        this.formulaVersion++;
        this.Columns.Clear();
        this.Rows.Clear();
        this.formulaCells.Clear();
        this.Charts.Clear();
        this.Merges = new List<MergeRange>();
        this.Images = new List<SheetImage>();
        this.DataBars = new List<SheetDataBar>();
        this.CellValueRules = new List<SheetCellValueRule>();
        this.AutoFilterRange = null;
        this.sourceHiddenRows.Clear();
        this.columnFilters.Clear();
        this.filterHeaderRow = 0;

        ColumnViewModel[] newColumns = new ColumnViewModel[columns];
        for (int c = 0; c < columns; c++)
        {
            newColumns[c] = new ColumnViewModel(c, this.DefaultColumnWidth);
        }

        this.Columns.AddRange(newColumns);

        RowViewModel[] newRows = new RowViewModel[rows];
        for (int r = 0; r < rows; r++)
        {
            RowViewModel row = new(r + 1, columns);
            for (int c = 0; c < row.Cells.Count; c++)
            {
                this.HookCell(row.Cells[c], row, c);
            }

            newRows[r] = row;
        }

        this.Rows.AddRange(newRows);
        this.ActiveCell = new CellAddress(0, 0);
    }

    public bool IsValid(CellAddress a) =>
        a.Row >= 0 && a.Row < this.RowCount && a.Column >= 0 && a.Column < this.ColumnCount;

    public CellViewModel? GetCell(CellAddress a) =>
        this.IsValid(a) ? this.Rows[a.Row].Cells[a.Column] : null;

    public void SetCell(CellAddress a, string text)
    {
        this.EnsureSize(a.Row + 1, a.Column + 1);
        this.Rows[a.Row].Cells[a.Column].Text = text;
    }

    // Clears every cell's content but keeps the current dimensions.
    public void Clear()
    {
        foreach (RowViewModel row in this.Rows)
        {
            foreach (CellViewModel cell in row.Cells)
            {
                cell.ApplyData(CellData.Plain(string.Empty));
            }
        }
    }

    // Highest row index (0-based) that contains any content, or -1 if the sheet is empty.
    public int LastNonEmptyRow()
    {
        for (int r = this.RowCount - 1; r >= 0; r--)
        {
            if (this.Rows[r].Cells.Any(c => !string.IsNullOrEmpty(c.Text)))
            {
                return r;
            }
        }

        return -1;
    }

    // Highest column index (0-based) that contains any content, or -1 if the sheet is empty.
    public int LastNonEmptyColumn()
    {
        for (int c = this.ColumnCount - 1; c >= 0; c--)
        {
            if (this.Rows.Any(row => c < row.Cells.Count && !string.IsNullOrEmpty(row.Cells[c].Text)))
            {
                return c;
            }
        }

        return -1;
    }

    // Sets the row count, growing or trimming trailing rows (used for dynamic scrolling).
    // Never trims below the last row containing data.
    public void SetRowCount(int rows)
    {
        rows = Math.Max(rows, this.LastNonEmptyRow() + 1);
        rows = Math.Max(rows, 1);
        if (rows > this.RowCount)
        {
            this.EnsureSize(rows, this.ColumnCount);
        }
        else
        {
            this.formulaCells.RemoveWhere(address => address.Row >= rows);
            while (this.RowCount > rows)
            {
                this.Rows.RemoveAt(this.RowCount - 1);
            }
        }
    }

    // Sets the column count, growing or trimming trailing columns (used for dynamic horizontal
    // scrolling). Never trims below the last column containing data.
    public void SetColumnCount(int columns)
    {
        columns = Math.Max(columns, this.LastNonEmptyColumn() + 1);
        columns = Math.Max(columns, 1);
        if (columns > this.ColumnCount)
        {
            this.EnsureSize(this.RowCount, columns);
        }
        else
        {
            this.formulaCells.RemoveWhere(address => address.Column >= columns);
            while (this.ColumnCount > columns)
            {
                this.Columns.RemoveAt(this.ColumnCount - 1);
                foreach (RowViewModel row in this.Rows)
                {
                    row.Cells.RemoveAt(row.Cells.Count - 1);
                }
            }
        }
    }

    // Grows the sheet so it has at least the requested rows/columns.
    public void EnsureSize(int rows, int columns)
    {
        while (this.ColumnCount < columns)
        {
            int index = this.ColumnCount;
            this.Columns.Add(new ColumnViewModel(index, this.DefaultColumnWidth));
            for (int rowIndex = 0; rowIndex < this.Rows.Count; rowIndex++)
            {
                CellViewModel cell = new();
                this.HookCell(cell, this.Rows[rowIndex], index);
                this.Rows[rowIndex].Cells.Add(cell);
            }
        }

        while (this.RowCount < rows)
        {
            RowViewModel row = new(this.RowCount + 1, this.ColumnCount);
            for (int column = 0; column < row.Cells.Count; column++)
            {
                this.HookCell(row.Cells[column], row, column);
            }

            this.Rows.Add(row);
        }
    }

    // Structural row operations keep the row objects (and their cell formatting) intact. The
    // HookCell callback uses RowViewModel.Number, so cells moved here continue to notify at their
    // new address. A single DataChanged is raised after the entire operation.
    public void InsertRows(int index, int count = 1)
    {
        this.CalculationVersion++;
        this.formulaVersion++;
        index = Math.Clamp(index, 0, this.RowCount);
        count = Math.Clamp(count, 1, 1000);
        for (int i = 0; i < count; i++)
        {
            RowViewModel row = new(index + i + 1, this.ColumnCount);
            for (int column = 0; column < row.Cells.Count; column++)
            {
                this.HookCell(row.Cells[column], row, column);
            }
            this.Rows.Insert(index + i, row);
        }
        this.RenumberRows(index);
        this.RemapFormulaCells(row => row >= index ? row + count : row);
        this.ShiftFormulaReferences(index, count, deleting: false);
        this.Merges = this.Merges.Select(merge => merge.Row >= index
            ? merge with { Row = merge.Row + count }
            : merge.Row + merge.RowSpan > index
                ? merge with { RowSpan = merge.RowSpan + count }
                : merge).ToList();
        this.ShiftRowAnchoredContent(index, count, deleting: false);
        if (this.columnFilters.Count > 0)
        {
            this.ApplyColumnFilters();
        }
        this.DataChanged?.Invoke();
    }

    public void DeleteRows(int index, int count = 1)
    {
        this.CalculationVersion++;
        this.formulaVersion++;
        if (this.RowCount == 0)
        {
            return;
        }
        index = Math.Clamp(index, 0, this.RowCount - 1);
        count = Math.Clamp(count, 1, this.RowCount - index);
        for (int i = 0; i < count; i++)
        {
            this.sourceHiddenRows.Remove(this.Rows[index]);
            this.Rows.RemoveAt(index);
        }
        if (this.Rows.Count == 0)
        {
            this.EnsureSize(1, Math.Max(1, this.ColumnCount));
        }
        this.RenumberRows(index);
        this.RemapFormulaCells(row => row < index ? row : row >= index + count ? row - count : -1);
        this.ShiftFormulaReferences(index, count, deleting: true);
        this.Merges = this.Merges.Select(merge =>
        {
            int first = merge.Row < index ? merge.Row : Math.Max(index, merge.Row - count);
            int last = merge.Row + merge.RowSpan - 1;
            int newLast = last < index ? last : last >= index + count ? last - count : index - 1;
            return new MergeRange(first, merge.Column, Math.Max(0, newLast - first + 1), merge.ColumnSpan);
        }).Where(static merge => merge.RowSpan > 0).ToList();
        this.ShiftRowAnchoredContent(index, count, deleting: true);
        if (this.columnFilters.Count > 0)
        {
            this.ApplyColumnFilters();
        }
        this.DataChanged?.Invoke();
    }

    public int LikelyHeaderRow()
    {
        if (this.AutoFilterRange is { } filter)
        {
            return Math.Clamp(filter.Row, 0, Math.Max(0, this.RowCount - 1));
        }
        int limit = Math.Min(this.RowCount, 50);
        for (int row = 0; row < limit; row++)
        {
            if (this.Rows[row].Cells.Count(static cell => !string.IsNullOrWhiteSpace(cell.Text)) >= 2)
            {
                return row;
            }
        }
        return 0;
    }

    private int DataEndRow(int headerRow)
    {
        if (this.AutoFilterRange is { } filter && headerRow == filter.Row)
        {
            return Math.Min(this.RowCount - 1, filter.Row + filter.RowSpan - 1);
        }
        int last = headerRow;
        for (int row = headerRow + 1; row < this.RowCount; row++)
        {
            if (this.Rows[row].Cells.All(static cell => string.IsNullOrWhiteSpace(cell.Text)))
            {
                break;
            }
            last = row;
        }
        return last;
    }

    public IReadOnlyList<string> ColumnValues(int column, int headerRow, int limit = 100)
    {
        if (column < 0 || column >= this.ColumnCount)
        {
            return Array.Empty<string>();
        }
        int last = this.DataEndRow(headerRow);
        return this.Rows.Skip(headerRow + 1).Take(Math.Max(0, last - headerRow))
            .Select(row => row.Cells[column].DisplayText)
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(static value => value, StringComparer.CurrentCultureIgnoreCase)
            .Take(limit).ToArray();
    }

    private readonly Dictionary<int, string> columnFilters = new();
    private int filterHeaderRow;

    public bool HasColumnFilter(int column) => this.columnFilters.ContainsKey(column);

    public void FilterColumn(int column, string? value, int headerRow)
    {
        this.filterHeaderRow = headerRow;
        if (value is null)
        {
            this.columnFilters.Remove(column);
        }
        else
        {
            this.columnFilters[column] = value;
        }
        this.ApplyColumnFilters();
        this.ViewChanged?.Invoke();
    }

    private void ApplyColumnFilters()
    {
        int last = this.DataEndRow(this.filterHeaderRow);
        for (int row = 0; row < this.RowCount; row++)
        {
            this.Rows[row].IsHidden = this.sourceHiddenRows.Contains(this.Rows[row]) ||
                row > this.filterHeaderRow && row <= last && this.columnFilters.Any(filter =>
                filter.Key < this.Rows[row].Cells.Count &&
                !string.Equals(this.Rows[row].Cells[filter.Key].DisplayText, filter.Value, StringComparison.CurrentCultureIgnoreCase));
        }
    }

    public void SortRows(int column, int headerRow, bool descending)
    {
        if (column < 0 || column >= this.ColumnCount || headerRow + 1 >= this.RowCount)
        {
            return;
        }
        if (this.AutoFilterRange is { } filter &&
            (column < filter.Column || column >= filter.Column + filter.ColumnSpan))
        {
            return;
        }
        int start = headerRow + 1;
        int last = this.DataEndRow(headerRow);
        if (last <= start)
        {
            return;
        }

        this.CalculationVersion++;
        this.formulaVersion++;
        this.engine.InvalidateCache();
        // Capture sort keys once. Reading/formating a formula cell during every comparison made
        // large worksheet sorts needlessly expensive, and comparing percent display strings was wrong.
        var ordered = this.Rows.Skip(start).Take(last - start + 1)
            .Select((row, original) => (Row: row, Original: original,
                Display: row.Cells[column].DisplayText, Number: row.Cells[column].NumericValue))
            .OrderBy(pair => pair, Comparer<(RowViewModel Row, int Original, string Display, double? Number)>.Create((left, right) =>
            {
                int result = left.Number is { } a && right.Number is { } b
                    ? a.CompareTo(b)
                    : StringComparer.CurrentCultureIgnoreCase.Compare(left.Display, right.Display);
                return result == 0 ? left.Original.CompareTo(right.Original) : descending ? -result : result;
            })).ToArray();
        CellAddress[] formulaAddresses = this.formulaCells.ToArray();
        int[] newIndexByOriginal = new int[ordered.Length];
        for (int i = 0; i < ordered.Length; i++)
        {
            newIndexByOriginal[ordered[i].Original] = start + i;
        }
        this.suspendRecalc = true;
        try
        {
            this.Rows.RemoveRange(start, ordered.Length);
            this.Rows.InsertRange(start, ordered.Select(static item => item.Row));
            this.RenumberRows(start);
            foreach (CellAddress address in formulaAddresses)
            {
                if (address.Row < start || address.Row > last)
                {
                    continue;
                }
                int newRow = newIndexByOriginal[address.Row - start];
                int delta = newRow - address.Row;
                if (delta != 0)
                {
                    CellViewModel cell = this.Rows[newRow].Cells[address.Column];
                    cell.RelocateFormula(ShiftFormulaForMovedRow(cell.Text, delta));
                }
            }
        }
        finally
        {
            this.suspendRecalc = false;
        }
        this.formulaCells.Clear();
        foreach (CellAddress address in formulaAddresses)
        {
            int row = address.Row >= start && address.Row <= last
                ? newIndexByOriginal[address.Row - start]
                : address.Row;
            this.formulaCells.Add(new CellAddress(row, address.Column));
        }
        if (this.columnFilters.Count > 0)
        {
            this.ApplyColumnFilters();
        }
        this.DataChanged?.Invoke();
    }

    // A sorted record keeps its row-relative formulas (for example =J3/I3 becomes =J5/I5
    // when that record moves from row 3 to row 5). Absolute rows and quoted text stay put.
    private static string ShiftFormulaForMovedRow(string formula, int delta)
    {
        static bool IdentifierChar(char value) => char.IsAsciiLetterOrDigit(value) || value is '_' or '.';

        bool inText = false;
        bool inSheetName = false;
        StringBuilder? rewritten = null;
        int copiedThrough = 0;
        for (int index = 0; index < formula.Length;)
        {
            char current = formula[index];
            if (current == '"' && !inSheetName)
            {
                if (inText && index + 1 < formula.Length && formula[index + 1] == '"')
                {
                    index += 2;
                    continue;
                }
                inText = !inText;
                index++;
                continue;
            }
            if (current == '\'' && !inText)
            {
                if (inSheetName && index + 1 < formula.Length && formula[index + 1] == '\'')
                {
                    index += 2;
                    continue;
                }
                inSheetName = !inSheetName;
                index++;
                continue;
            }
            if (inText || inSheetName || current is not ('$' or >= 'A' and <= 'Z' or >= 'a' and <= 'z') ||
                index > 0 && IdentifierChar(formula[index - 1]))
            {
                index++;
                continue;
            }

            int start = index;
            int next = current == '$' ? index + 1 : index;
            int letterStart = next;
            while (next < formula.Length && char.IsAsciiLetter(formula[next]))
            {
                next++;
            }
            int letters = next - letterStart;
            if (letters is < 1 or > 3)
            {
                index = start + 1;
                continue;
            }
            bool absoluteRow = next < formula.Length && formula[next] == '$';
            if (absoluteRow)
            {
                next++;
            }
            int rowStart = next;
            if (next >= formula.Length || formula[next] is < '1' or > '9')
            {
                index = start + 1;
                continue;
            }
            do
            {
                next++;
            }
            while (next < formula.Length && char.IsAsciiDigit(formula[next]));
            if (next < formula.Length && (IdentifierChar(formula[next]) || formula[next] == '(') ||
                absoluteRow ||
                !int.TryParse(formula.AsSpan(rowStart, next - rowStart), NumberStyles.None,
                    CultureInfo.InvariantCulture, out int row))
            {
                index = next;
                continue;
            }

            rewritten ??= new StringBuilder(formula.Length + 8);
            rewritten.Append(formula, copiedThrough, rowStart - copiedThrough);
            rewritten.Append(Math.Max(1, row + delta).ToString(CultureInfo.InvariantCulture));
            copiedThrough = next;
            index = next;
        }
        if (rewritten is null)
        {
            return formula;
        }
        rewritten.Append(formula, copiedThrough, formula.Length - copiedThrough);
        return rewritten.ToString();
    }

    private void RenumberRows(int start)
    {
        for (int row = Math.Max(0, start); row < this.RowCount; row++)
        {
            this.Rows[row].Number = row + 1;
        }
        this.ActiveCell = new CellAddress(Math.Clamp(this.ActiveCell.Row, 0, this.RowCount - 1),
            Math.Clamp(this.ActiveCell.Column, 0, this.ColumnCount - 1));
    }

    private void RemapFormulaCells(Func<int, int> mapRow)
    {
        CellAddress[] addresses = this.formulaCells.ToArray();
        this.formulaCells.Clear();
        foreach (CellAddress address in addresses)
        {
            int row = mapRow(address.Row);
            if (row >= 0)
            {
                this.formulaCells.Add(new CellAddress(row, address.Column));
            }
        }
    }

    private void ShiftRowAnchoredContent(int index, int count, bool deleting)
    {
        int Shift(int row) => row < index ? row : deleting
            ? Math.Max(index, row - count)
            : row + count;

        CellRange? ShiftRange(CellRange range)
        {
            int first = range.Row;
            int end = range.Row + range.RowSpan;
            if (deleting)
            {
                first = first < index ? first : Math.Max(index, first - count);
                end = end <= index ? end : Math.Max(index, end - count);
            }
            else if (first >= index)
            {
                first += count;
                end += count;
            }
            else if (end > index)
            {
                end += count;
            }
            return end > first ? range with { Row = first, RowSpan = end - first } : null;
        }

        if (this.AutoFilterRange is { } autoFilter)
        {
            this.AutoFilterRange = ShiftRange(autoFilter);
        }

        this.DataBars = this.DataBars.Select(bar => (Bar: bar, Range: ShiftRange(bar.Range)))
            .Where(static item => item.Range is not null)
            .Select(static item => new SheetDataBar
            {
                Range = item.Range!.Value,
                Min = item.Bar.Min,
                Max = item.Bar.Max,
                Color = item.Bar.Color,
                ShowValue = item.Bar.ShowValue,
                Gradient = item.Bar.Gradient,
            }).ToList();
        this.CellValueRules = this.CellValueRules.Select(rule => (Rule: rule, Range: ShiftRange(rule.Range)))
            .Where(static item => item.Range is not null)
            .Select(static item => new SheetCellValueRule
            {
                Range = item.Range!.Value,
                Comparison = item.Rule.Comparison,
                Threshold = item.Rule.Threshold,
                Foreground = item.Rule.Foreground,
                Bold = item.Rule.Bold,
                Priority = item.Rule.Priority,
            }).ToList();
        this.Images = this.Images.Select(image => new SheetImage
        {
            Row = Shift(image.Row),
            Column = image.Column,
            OffsetX = image.OffsetX,
            OffsetY = image.OffsetY,
            Width = image.Width,
            Height = image.Height,
            Data = image.Data,
        }).ToList();
        foreach (SheetChart chart in this.Charts)
        {
            chart.AnchorRow = Shift(chart.AnchorRow);
            if (chart.ToRow is { } toRow)
            {
                chart.ToRow = Shift(toRow);
            }
        }
        if (this.columnFilters.Count > 0)
        {
            this.filterHeaderRow = Shift(this.filterHeaderRow);
        }
    }

    private static readonly Regex A1Reference = new(
        @"(?<![A-Za-z0-9_.])(?<column>\$?[A-Za-z]{1,3})(?<row>\$?[1-9][0-9]*)(?![A-Za-z0-9_.])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private void ShiftFormulaReferences(int index, int count, bool deleting)
    {
        this.suspendRecalc = true;
        try
        {
            foreach (CellAddress address in this.formulaCells.ToArray())
            {
                CellViewModel? cell = this.GetCell(address);
                if (cell is not null && FormulaEngine.IsFormula(cell.Text))
                {
                    cell.Text = this.RewriteRowReferences(cell.Text, index, count, deleting);
                }
            }
            this.names = this.names.ToDictionary(
                static pair => pair.Key,
                pair => this.RewriteRowReferences(pair.Value, index, count, deleting),
                StringComparer.OrdinalIgnoreCase);
        }
        finally
        {
            this.suspendRecalc = false;
        }
    }

    private string RewriteRowReferences(string formula, int index, int count, bool deleting) =>
        A1Reference.Replace(formula, match =>
        {
            // Quoted string literals (including doubled Excel quotes) are not references.
            bool inString = false;
            for (int i = 0; i < match.Index; i++)
            {
                if (formula[i] != '"')
                {
                    continue;
                }
                if (inString && i + 1 < match.Index && formula[i + 1] == '"')
                {
                    i++;
                }
                else
                {
                    inString = !inString;
                }
            }
            if (inString)
            {
                return match.Value;
            }

            // Do not rewrite a reference explicitly qualified with another sheet name.
            if (match.Index > 0 && formula[match.Index - 1] == '!')
            {
                int end = match.Index - 1;
                int start = end - 1;
                if (start >= 0 && formula[start] == '\'')
                {
                    start = formula.LastIndexOf('\'', start - 1);
                    if (start < 0)
                    {
                        return match.Value;
                    }
                    start++;
                }
                else
                {
                    while (start >= 0 && (char.IsLetterOrDigit(formula[start]) || formula[start] == '_'))
                    {
                        start--;
                    }
                    start++;
                }
                string referencedSheet = formula[start..end].Trim('\'');
                if (!referencedSheet.Equals(this.Name, StringComparison.OrdinalIgnoreCase))
                {
                    return match.Value;
                }
            }

            string rowToken = match.Groups["row"].Value;
            if (!int.TryParse(rowToken.TrimStart('$'), out int oneBasedRow))
            {
                return match.Value;
            }
            int row = oneBasedRow - 1;
            if (row < index)
            {
                return match.Value;
            }
            if (deleting && row < index + count)
            {
                return "#REF!";
            }
            int changed = deleting ? row - count : row + count;
            return match.Groups["column"].Value + (rowToken.StartsWith('$') ? "$" : string.Empty) + (changed + 1);
        });

    private void ReindexFormulaCells()
    {
        this.CalculationVersion++;
        this.formulaVersion++;
        this.formulaCells.Clear();
        for (int row = 0; row < this.RowCount; row++)
        {
            for (int column = 0; column < this.ColumnCount; column++)
            {
                if (FormulaEngine.IsFormula(this.Rows[row].Cells[column].Text))
                {
                    this.formulaCells.Add(new CellAddress(row, column));
                }
            }
        }
    }

    // Replaces all content from a rectangular set of plain-text rows (e.g. parsed TSV/CSV).
    public void LoadRows(IReadOnlyList<IReadOnlyList<string>> data)
    {
        int rows = Math.Max(data.Count, 1);
        int cols = Math.Max(data.Count == 0 ? 1 : data.Max(r => r.Count), 1);
        this.Initialize(rows, cols);
        this.suspendRecalc = true;
        for (int r = 0; r < data.Count; r++)
        {
            for (int c = 0; c < data[r].Count; c++)
            {
                this.Rows[r].Cells[c].Text = data[r][c];
            }
        }

        this.suspendRecalc = false;
        this.RequestRecalculation();
    }

    // Projects the sheet back to plain-text rows for value-only serialization.
    public List<List<string>> ToRows() =>
        this.Rows.Select(r => r.Cells.Select(c => c.Text).ToList()).ToList();

    // Replaces all content from rows of CellData (value + formatting).
    public void LoadCellData(IReadOnlyList<IReadOnlyList<CellData>> data, bool recalculate = true)
    {
        int rows = Math.Max(data.Count, 1);
        int cols = Math.Max(data.Count == 0 ? 1 : data.Max(r => r.Count), 1);
        this.Initialize(rows, cols);
        this.suspendRecalc = true;
        for (int r = 0; r < data.Count; r++)
        {
            for (int c = 0; c < data[r].Count; c++)
            {
                this.Rows[r].Cells[c].ApplyData(data[r][c]);
            }
        }

        this.suspendRecalc = false;
        if (recalculate)
        {
            this.RequestRecalculation();
        }
    }

    // Projects the sheet to rows of CellData (value + formatting).
    public List<List<CellData>> ToCellData() =>
        this.Rows.Select(r => r.Cells.Select(c => c.ToData()).ToList()).ToList();

    // ----- Undo / redo (snapshot-based; covers cell values + formatting, not charts/merges) -----

    private const int MaxUndoDepth = 100;
    private readonly List<List<List<CellData>>> undoStack = new();
    private readonly List<List<List<CellData>>> redoStack = new();

    public bool CanUndo => this.undoStack.Count > 0;

    public bool CanRedo => this.redoStack.Count > 0;

    // Raised when the undo/redo availability changes (so the toolbar can enable/disable its buttons).
    public event Action? UndoStateChanged;

    // Records the current cell state as an undo point. Call BEFORE a mutation. No-ops if the state is
    // unchanged from the most recent checkpoint, so repeated calls don't pile up identical entries.
    public void Checkpoint()
    {
        List<List<CellData>> snapshot = this.ToCellData();
        if (this.undoStack.Count > 0 && SameCells(this.undoStack[^1], snapshot))
        {
            return;
        }

        this.undoStack.Add(snapshot);
        if (this.undoStack.Count > MaxUndoDepth)
        {
            this.undoStack.RemoveAt(0);
        }

        this.redoStack.Clear();
        this.UndoStateChanged?.Invoke();
    }

    public void Undo()
    {
        if (this.undoStack.Count == 0)
        {
            return;
        }

        this.redoStack.Add(this.ToCellData());
        List<List<CellData>> snapshot = this.undoStack[^1];
        this.undoStack.RemoveAt(this.undoStack.Count - 1);
        this.RestoreCells(snapshot);
        this.UndoStateChanged?.Invoke();
    }

    public void Redo()
    {
        if (this.redoStack.Count == 0)
        {
            return;
        }

        this.undoStack.Add(this.ToCellData());
        List<List<CellData>> snapshot = this.redoStack[^1];
        this.redoStack.RemoveAt(this.redoStack.Count - 1);
        this.RestoreCells(snapshot);
        this.UndoStateChanged?.Invoke();
    }

    // Restores a cell snapshot in place — resizes to the snapshot's dimensions and re-applies each
    // cell — WITHOUT touching charts/merges/images/names (so undo of an edit keeps those intact).
    private void RestoreCells(List<List<CellData>> snapshot)
    {
        int rows = Math.Max(snapshot.Count, 1);
        int cols = Math.Max(snapshot.Count == 0 ? 1 : snapshot.Max(r => r.Count), 1);
        this.EnsureSize(rows, cols);
        while (this.RowCount > rows)
        {
            this.Rows.RemoveAt(this.RowCount - 1);
        }

        while (this.ColumnCount > cols)
        {
            this.Columns.RemoveAt(this.ColumnCount - 1);
            foreach (RowViewModel row in this.Rows)
            {
                row.Cells.RemoveAt(row.Cells.Count - 1);
            }
        }

        this.suspendRecalc = true;
        for (int r = 0; r < snapshot.Count; r++)
        {
            for (int c = 0; c < snapshot[r].Count && c < this.Rows[r].Cells.Count; c++)
            {
                this.Rows[r].Cells[c].ApplyData(snapshot[r][c]);
            }
        }

        this.suspendRecalc = false;
        this.RequestRecalculation();
    }

    private static bool SameCells(List<List<CellData>> a, List<List<CellData>> b)
    {
        if (a.Count != b.Count)
        {
            return false;
        }

        for (int r = 0; r < a.Count; r++)
        {
            if (a[r].Count != b[r].Count || !a[r].SequenceEqual(b[r]))
            {
                return false;
            }
        }

        return true;
    }
}
