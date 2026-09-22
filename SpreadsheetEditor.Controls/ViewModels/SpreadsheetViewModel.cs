namespace SpreadsheetEditor.Controls.ViewModels;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

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
    private string name = "Sheet1";

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

    // Charts embedded in the sheet. The grid renders these as live overlays.
    public AvaloniaList<SheetChart> Charts { get; } = new();

    // Raised after a recalculation so chart overlays can refresh their data.
    public event Action? DataChanged;

    private readonly FormulaEngine engine;

    private bool suspendRecalc;

    private Dictionary<string, string> names = new(StringComparer.OrdinalIgnoreCase);

    public SpreadsheetViewModel() => this.engine = new FormulaEngine(this.ReadRaw, this.ResolveName);

    // Defined names mapped to their formula text (e.g. NetIncome -> "O36",
    // TotalMonthlyIncome -> "SUM('Sheet'!$C$7:$C$10)") so formulas like =NetIncome resolve.
    public void SetNames(IReadOnlyDictionary<string, string> names) =>
        this.names = new Dictionary<string, string>(names, StringComparer.OrdinalIgnoreCase);

    // The defined names (name -> RefersTo formula/range), so a save can persist them back to the
    // .xlsx — otherwise formulas that reference a name become #NAME? on the next open.
    public IReadOnlyDictionary<string, string> Names => this.names;

    private string? ResolveName(string name) => this.names.TryGetValue(name, out string? f) ? f : null;

    public SpreadsheetViewModel(int rows, int columns) : this() => this.Initialize(rows, columns);

    private string ReadRaw(CellAddress a) => this.IsValid(a) ? this.Rows[a.Row].Cells[a.Column].Text : string.Empty;

    // Runs a bulk edit with recalculation suspended, then recalculates once at the end.
    public void RunBatch(Action action)
    {
        this.suspendRecalc = true;
        try
        {
            action();
        }
        finally
        {
            this.suspendRecalc = false;
        }

        this.Recalculate();
    }

    // Recomputes every formula cell from the current cell contents.
    public void Recalculate()
    {
        List<CellAddress> addresses = new(this.RowCount * this.ColumnCount);
        for (int r = 0; r < this.RowCount; r++)
        {
            for (int c = 0; c < this.Rows[r].Cells.Count; c++)
            {
                addresses.Add(new CellAddress(r, c));
            }
        }

        this.engine.Recalculate(addresses, (addr, value) =>
        {
            CellViewModel cell = this.Rows[addr.Row].Cells[addr.Column];
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
        this.DataChanged?.Invoke();
    }

    // Resizes a chart to an explicit pixel size (keeping its anchor) and re-lays-out the overlay.
    public void ResizeChart(SheetChart chart, double width, double height)
    {
        chart.ResizeTo(width, height);
        this.DataChanged?.Invoke();
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
            string seriesTitle = header ? this.GetCell(new CellAddress(range.Row, c))?.DisplayText ?? "" : $"Series {c - valueColStart + 1}";
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
    private void HookCell(CellViewModel cell)
    {
        cell.PropertyChanged += (_, e) =>
        {
            if (e.PropertyName == nameof(CellViewModel.Text) && !this.suspendRecalc)
            {
                this.Recalculate();
            }
        };
    }

    // (Re)builds an empty sheet of the given size.
    public void Initialize(int rows, int columns)
    {
        this.Columns.Clear();
        this.Rows.Clear();
        this.Charts.Clear();
        this.Merges = new List<MergeRange>();
        this.Images = new List<SheetImage>();
        this.DataBars = new List<SheetDataBar>();

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
            foreach (CellViewModel cell in row.Cells)
            {
                this.HookCell(cell);
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
            foreach (RowViewModel row in this.Rows)
            {
                CellViewModel cell = new();
                this.HookCell(cell);
                row.Cells.Add(cell);
            }
        }

        while (this.RowCount < rows)
        {
            RowViewModel row = new(this.RowCount + 1, this.ColumnCount);
            foreach (CellViewModel cell in row.Cells)
            {
                this.HookCell(cell);
            }

            this.Rows.Add(row);
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
        this.Recalculate();
    }

    // Projects the sheet back to plain-text rows for value-only serialization.
    public List<List<string>> ToRows() =>
        this.Rows.Select(r => r.Cells.Select(c => c.Text).ToList()).ToList();

    // Replaces all content from rows of CellData (value + formatting).
    public void LoadCellData(IReadOnlyList<IReadOnlyList<CellData>> data)
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
        this.Recalculate();
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
        this.Recalculate();
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