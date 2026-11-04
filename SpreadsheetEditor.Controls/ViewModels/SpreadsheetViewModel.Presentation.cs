namespace SpreadsheetEditor.Controls.ViewModels;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using global::SpreadsheetEditor.Controls.Models;

public partial class SpreadsheetViewModel
{
    private readonly HashSet<int> sourceHiddenColumns = new();
    private IReadOnlyList<double> sourceColumnWidths = Array.Empty<double>();

    public void SetColumnWidths(IReadOnlyList<double> widths)
    {
        this.sourceColumnWidths = widths.ToArray();
        for (int c = 0; c < Math.Min(this.ColumnCount, widths.Count); c++)
            if (widths[c] > 0) this.Columns[c].Width = widths[c];
    }

    public IReadOnlyList<double> ColumnWidthSnapshot()
    {
        double[] widths = new double[Math.Max(this.ColumnCount, this.sourceColumnWidths.Count)];
        for (int c = 0; c < widths.Length; c++)
            widths[c] = c < this.ColumnCount ? this.Columns[c].Width : this.sourceColumnWidths[c];
        return widths;
    }

    public void SetHiddenColumns(IReadOnlyList<int> columns)
    {
        this.sourceHiddenColumns.Clear();
        this.sourceHiddenColumns.UnionWith(columns.Where(static index => index >= 0));
        for (int index = 0; index < this.Columns.Count; index++)
            this.Columns[index].IsHidden = this.sourceHiddenColumns.Contains(index);
    }

    public IReadOnlyList<int> HiddenColumnSnapshot() => this.Columns
        .Where(static column => column.IsHidden).Select(static column => column.Index)
        .Concat(this.sourceHiddenColumns.Where(index => index >= this.ColumnCount)).Order().ToArray();

    public string Visibility { get; set; } = "visible";
    private string? protectionXml;
    public string? ProtectionXml
    {
        get => this.protectionXml;
        set
        {
            this.protectionXml = value;
            this.IsProtected = value is not null && XElement.Parse(value).Attribute("sheet")?.Value is "1" or "true";
        }
    }
    public bool IsProtected { get; private set; }
    public IReadOnlyList<SheetOutlineGroup> RowGroups { get; set; } = Array.Empty<SheetOutlineGroup>();
    public IReadOnlyList<SheetOutlineGroup> ColumnGroups { get; set; } = Array.Empty<SheetOutlineGroup>();

    public bool CanEdit(CellAddress address) => !this.IsProtected || this.GetCell(address)?.Appearance?.Locked == false;

    public bool CanEdit(CellRange range)
    {
        if (!this.IsProtected) return true;
        for (int row = range.Row; row < range.Row + range.RowSpan; row++)
            for (int column = range.Column; column < range.Column + range.ColumnSpan; column++)
                if (!this.CanEdit(new CellAddress(row, column))) return false;
        return true;
    }

    public void ToggleRowGroup(SheetOutlineGroup group)
    {
        bool collapse = !group.Collapsed;
        this.RowGroups = this.RowGroups.Select(g => g == group ? g with { Collapsed = collapse } : g).ToArray();
        for (int r = group.Start; r < Math.Min(this.RowCount, group.Start + group.Count); r++)
        {
            bool hidden = this.RowGroups.Any(g => g.Collapsed && r >= g.Start && r < g.Start + g.Count);
            if (hidden) this.sourceHiddenRows.Add(this.Rows[r]); else this.sourceHiddenRows.Remove(this.Rows[r]);
        }
        this.ApplyColumnFilters();
        this.NotifyMetadataChanged();
    }
}
