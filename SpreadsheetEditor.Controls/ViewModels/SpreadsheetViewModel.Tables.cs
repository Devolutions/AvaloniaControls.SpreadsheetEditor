namespace SpreadsheetEditor.Controls.ViewModels;

using System;
using System.Collections.Generic;
using System.Linq;

using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Models;

public partial class SpreadsheetViewModel
{
    public IReadOnlyList<SheetTable> Tables { get; private set; } = Array.Empty<SheetTable>();

    public void SetTables(IReadOnlyList<SheetTable> tables) => this.Tables = tables.ToArray();

    private IEnumerable<CellRange> FilterRanges => this.Tables
        .Where(table => table.ShowHeaderRow && table.AutoFilterRange is not null)
        .Select(table => table.AutoFilterRange!.Value)
        .Concat(this.AutoFilterRange is { } range ? new[] { range } : Array.Empty<CellRange>());

    public CellRange? FilterHeaderRange(int row, int column)
    {
        foreach (CellRange range in this.FilterRanges)
            if (row == range.Row && column >= range.Column && column < range.Column + range.ColumnSpan) return range;
        return null;
    }

    public int ColumnHeaderRow(int column)
    {
        CellRange[] ranges = this.FilterRanges.Where(range => column >= range.Column && column < range.Column + range.ColumnSpan).ToArray();
        foreach (CellRange range in ranges)
            if (this.ActiveCell.Row >= range.Row && this.ActiveCell.Row < range.Row + range.RowSpan) return range.Row;
        return ranges.Length > 0 ? ranges[0].Row : this.LikelyHeaderRow();
    }

    public int LikelyHeaderRow()
    {
        if (this.FilterRanges.FirstOrDefault() is { RowSpan: > 0 } filter)
            return Math.Clamp(filter.Row, 0, Math.Max(0, this.RowCount - 1));
        for (int row = 0; row < Math.Min(this.RowCount, 50); row++)
            if (this.Rows[row].Cells.Count(cell => !string.IsNullOrWhiteSpace(cell.Text)) >= 2) return row;
        return 0;
    }

    private int DataEndRow(int headerRow, int column)
    {
        if (this.FilterHeaderRange(headerRow, column) is { } range)
            return Math.Min(this.RowCount - 1, range.Row + range.RowSpan - 1);
        int last = headerRow;
        for (int row = headerRow + 1; row < this.RowCount; row++)
        {
            if (this.Rows[row].Cells.All(cell => string.IsNullOrWhiteSpace(cell.Text))) break;
            last = row;
        }
        return last;
    }

    public IReadOnlyList<string> ColumnValues(int column, int headerRow, int limit = 100)
    {
        if (!this.IsFilterColumn(column, headerRow)) return Array.Empty<string>();
        int last = this.DataEndRow(headerRow, column);
        return this.Rows.Skip(headerRow + 1).Take(Math.Max(0, last - headerRow))
            .Select(row => row.Cells[column].DisplayText)
            .Distinct(StringComparer.CurrentCultureIgnoreCase)
            .OrderBy(value => value, StringComparer.CurrentCultureIgnoreCase).Take(limit).ToArray();
    }

    private bool IsFilterColumn(int column, int headerRow) => column >= 0 && column < this.ColumnCount &&
        headerRow >= 0 && headerRow < this.RowCount &&
        (!this.FilterRanges.Any(range => range.Row == headerRow) || this.FilterHeaderRange(headerRow, column) is not null);

    private readonly Dictionary<(int HeaderRow, int Column), HashSet<string>> columnFilters = new();

    public bool HasColumnFilter(int column, int? headerRow = null) => this.columnFilters.Keys.Any(key =>
        key.Column == column && (headerRow is null || key.HeaderRow == headerRow));

    public IReadOnlyList<SheetColumnFilter> ColumnFilterSnapshot() => this.columnFilters.Select(pair =>
        new SheetColumnFilter(pair.Key.HeaderRow, pair.Key.Column, pair.Value.ToArray())).ToArray();

    public void SetColumnFilters(IReadOnlyList<SheetColumnFilter> filters)
    {
        this.columnFilters.Clear();
        foreach (SheetColumnFilter filter in filters)
            if (this.IsFilterColumn(filter.Column, filter.HeaderRow))
                this.columnFilters[(filter.HeaderRow, filter.Column)] = new HashSet<string>(filter.Values, StringComparer.CurrentCultureIgnoreCase);
        this.ApplyColumnFilters();
    }

    public void FilterColumn(int column, string? value, int headerRow)
    {
        if (!this.IsFilterColumn(column, headerRow)) return;
        if (value is null) this.columnFilters.Remove((headerRow, column));
        else this.columnFilters[(headerRow, column)] = new HashSet<string>(new[] { value }, StringComparer.CurrentCultureIgnoreCase);
        this.ApplyColumnFilters();
        this.ViewChanged?.Invoke();
    }

    public void ClearTableFilters(int column, int headerRow)
    {
        if (this.FilterHeaderRange(headerRow, column) is not { } range) return;
        foreach (var key in this.columnFilters.Keys.Where(key => key.HeaderRow == range.Row &&
            key.Column >= range.Column && key.Column < range.Column + range.ColumnSpan).ToArray()) this.columnFilters.Remove(key);
        this.ApplyColumnFilters();
        this.ViewChanged?.Invoke();
    }

    private bool IsFilteredOut(int row) => this.columnFilters.Any(filter => row > filter.Key.HeaderRow &&
        row <= this.DataEndRow(filter.Key.HeaderRow, filter.Key.Column) &&
        !filter.Value.Contains(this.Rows[row].Cells[filter.Key.Column].DisplayText));

    private void ApplyColumnFilters()
    {
        var filters = this.columnFilters.Select(filter => (filter.Key.HeaderRow, filter.Key.Column,
            Last: this.DataEndRow(filter.Key.HeaderRow, filter.Key.Column), Values: filter.Value)).ToArray();
        for (int row = 0; row < this.RowCount; row++)
            this.Rows[row].IsHidden = this.sourceHiddenRows.Contains(this.Rows[row]) || filters.Any(filter =>
                row > filter.HeaderRow && row <= filter.Last && !filter.Values.Contains(this.Rows[row].Cells[filter.Column].DisplayText));
    }

    private void SortTableCells(CellRange range, int column, bool descending)
    {
        int first = range.Row + 1, last = Math.Min(this.RowCount - 1, range.Row + range.RowSpan - 1);
        int endColumn = Math.Min(this.ColumnCount, range.Column + range.ColumnSpan);
        if (last <= first || range.Column < 0 || endColumn <= range.Column) return;
        // Sort only the rectangular table body. Moving whole RowViewModels also moved unrelated
        // cells beside a table (notes, another table, confirmations, etc.).
        var ordered = this.Rows.Skip(first).Take(last - first + 1).Select((row, index) => new
        {
            OriginalRow = first + index,
            Key = row.Cells[column].DisplayText,
            Number = row.Cells[column].NumericValue,
            Cells = row.Cells.Skip(range.Column).Take(endColumn - range.Column).Select(cell => cell.ToData()).ToArray(),
            Runs = row.Cells.Skip(range.Column).Take(endColumn - range.Column).Select(cell => cell.Runs?.ToList()).ToArray(),
        }).ToArray();
        Array.Sort(ordered, (left, right) =>
        {
            int comparison = left.Number is { } a && right.Number is { } b ? a.CompareTo(b) :
                StringComparer.CurrentCultureIgnoreCase.Compare(left.Key, right.Key);
            return comparison == 0 ? left.OriginalRow.CompareTo(right.OriginalRow) : descending ? -comparison : comparison;
        });
        this.CalculationVersion++;
        this.formulaVersion++;
        this.engine.InvalidateCache();
        bool wasSuspended = this.suspendRecalc;
        this.suspendRecalc = true;
        try
        {
            for (int index = 0; index < ordered.Length; index++)
            {
                int row = first + index;
                for (int offset = 0; offset < ordered[index].Cells.Length; offset++)
                {
                    CellData data = ordered[index].Cells[offset];
                    if (FormulaEngine.IsFormula(data.Text) && row != ordered[index].OriginalRow)
                        data = data with { Text = ShiftFormulaForMovedRow(data.Text, row - ordered[index].OriginalRow) };
                    CellViewModel destination = this.Rows[row].Cells[range.Column + offset];
                    destination.ApplyData(data);
                    if (!destination.IsFormula && ordered[index].Runs[offset] is { } runs) destination.SetRichRuns(runs);
                }
            }
        }
        finally { this.suspendRecalc = wasSuspended; }
        if (!wasSuspended) this.batchedChangedCells.Clear();
        this.ApplyColumnFilters();
        this.DataChanged?.Invoke();
    }
}
