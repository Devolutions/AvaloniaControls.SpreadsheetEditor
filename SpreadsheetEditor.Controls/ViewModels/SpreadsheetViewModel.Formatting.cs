namespace SpreadsheetEditor.Controls.ViewModels;

using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Models;

public partial class SpreadsheetViewModel
{
    public CellRange ExpandFormattingRange(CellRange range)
    {
        bool expanded;
        do
        {
            expanded = false;
            foreach (MergeRange merge in this.Merges)
            {
                if (!Intersects(range, merge)) continue;
                int row = Math.Min(range.Row, merge.Row);
                int column = Math.Min(range.Column, merge.Column);
                CellRange union = new(row, column,
                    Math.Max(range.Row + range.RowSpan, merge.Row + merge.RowSpan) - row,
                    Math.Max(range.Column + range.ColumnSpan, merge.Column + merge.ColumnSpan) - column);
                if (union != range) { range = union; expanded = true; }
            }
        } while (expanded);
        return range;
    }

    private bool CanFormatRange(CellRange range) => !this.IsProtected && range.Row >= 0 && range.Column >= 0
        && range.RowSpan > 0 && range.ColumnSpan > 0
        && (long)range.Row + range.RowSpan <= this.RowCount
        && (long)range.Column + range.ColumnSpan <= this.ColumnCount;

    public bool CanMergeCells(CellRange range) => this.CanFormatRange(range) && this.CanFormatRange(range = this.ExpandFormattingRange(range))
        && !this.Tables.Any(table => range.Row < table.Range.Row + table.Range.RowSpan && range.Row + range.RowSpan > table.Range.Row
            && range.Column < table.Range.Column + table.Range.ColumnSpan && range.Column + range.ColumnSpan > table.Range.Column);

    public bool MergeWouldDiscardData(CellRange range, CellMergeMode mode)
    {
        if (!this.CanFormatRange(range)) return false;
        range = this.ExpandFormattingRange(range);
        if (!this.CanFormatRange(range)) return false;
        for (int row = range.Row; row < range.Row + range.RowSpan; row++)
            for (int column = range.Column; column < range.Column + range.ColumnSpan; column++)
                if ((column != range.Column || mode != CellMergeMode.Across && row != range.Row)
                    && !string.IsNullOrEmpty(this.Rows[row].Cells[column].Text)) return true;
        return false;
    }

    public bool MergeCells(CellRange range, CellMergeMode mode = CellMergeMode.Merge, bool discardContents = false)
    {
        if (!this.CanFormatRange(range)) return false;
        range = this.ExpandFormattingRange(range);
        if (!Enum.IsDefined(mode) || !this.CanMergeCells(range)
            || range.ColumnSpan == 1 && (mode == CellMergeMode.Across || range.RowSpan == 1)) return false;
        bool discard = this.MergeWouldDiscardData(range, mode);
        if (discard && !discardContents) return false;
        this.Checkpoint();
        List<MergeRange> merges = this.Merges.Where(merge => !Intersects(range, merge)).ToList();
        if (mode == CellMergeMode.Across)
        {
            for (int row = range.Row; row < range.Row + range.RowSpan; row++)
                merges.Add(new MergeRange(row, range.Column, 1, range.ColumnSpan));
        }
        else merges.Add(new MergeRange(range.Row, range.Column, range.RowSpan, range.ColumnSpan));
        this.SetMerges(merges);
        if (mode == CellMergeMode.Center) this.Rows[range.Row].Cells[range.Column].TextAlignment = TextAlignment.Center;
        if (discard)
        {
            this.RunBatch(() =>
            {
                for (int row = range.Row; row < range.Row + range.RowSpan; row++)
                    for (int column = range.Column; column < range.Column + range.ColumnSpan; column++)
                        if (column != range.Column || mode != CellMergeMode.Across && row != range.Row)
                            this.Rows[row].Cells[column].Text = string.Empty;
            });
        }
        this.NotifyMetadataChanged();
        return true;
    }

    public bool UnmergeCells(CellRange range)
    {
        if (!this.CanFormatRange(range)) return false;
        range = this.ExpandFormattingRange(range);
        if (!this.CanFormatRange(range) || !this.Merges.Any(merge => Intersects(range, merge))) return false;
        this.Checkpoint();
        this.SetMerges(this.Merges.Where(merge => !Intersects(range, merge)).ToArray());
        this.NotifyMetadataChanged();
        return true;
    }

    public bool ApplyBorders(CellRange range, CellBorderPreset preset, CellEdge edge)
    {
        if (!this.CanFormatRange(range)) return false;
        range = this.ExpandFormattingRange(range);
        if (!this.CanFormatRange(range) || !Enum.IsDefined(preset)) return false;
        this.Checkpoint();
        this.SetBorders(range, preset, edge);
        this.NotifyMetadataChanged();
        return true;
    }

    [Flags]
    private enum BorderSides { Left = 1, Right = 2, Top = 4, Bottom = 8, All = 15 }

    private IEnumerable<(CellViewModel Cell, BorderSides Sides)> BorderTargets(CellRange range)
    {
        for (int row = range.Row; row < range.Row + range.RowSpan; row++)
        {
            for (int column = range.Column; column < range.Column + range.ColumnSpan; column++)
                yield return (this.Rows[row].Cells[column], BorderSides.All);
            if (range.Column > 0) yield return (this.Rows[row].Cells[range.Column - 1], BorderSides.Right);
            if (range.Column + range.ColumnSpan < this.ColumnCount)
                yield return (this.Rows[row].Cells[range.Column + range.ColumnSpan], BorderSides.Left);
        }
        for (int column = range.Column; column < range.Column + range.ColumnSpan; column++)
        {
            if (range.Row > 0) yield return (this.Rows[range.Row - 1].Cells[column], BorderSides.Bottom);
            if (range.Row + range.RowSpan < this.RowCount)
                yield return (this.Rows[range.Row + range.RowSpan].Cells[column], BorderSides.Top);
        }
    }

    internal IEnumerable<CellEdge> ExistingBorderEdges(CellRange range)
    {
        if (!this.CanFormatRange(range)) yield break;
        range = this.ExpandFormattingRange(range);
        if (!this.CanFormatRange(range)) yield break;
        foreach ((CellViewModel cell, BorderSides sides) in this.BorderTargets(range))
        {
            if (cell.Appearance is not { } appearance) continue;
            if ((sides & BorderSides.Left) != 0 && appearance.Left is { } left) yield return left;
            if ((sides & BorderSides.Right) != 0 && appearance.Right is { } right) yield return right;
            if ((sides & BorderSides.Top) != 0 && appearance.Top is { } top) yield return top;
            if ((sides & BorderSides.Bottom) != 0 && appearance.Bottom is { } bottom) yield return bottom;
        }
    }

    public bool UpdateBorderFormatting(CellRange range, Color? color = null, string? style = null)
    {
        if (!this.SetExistingBorderFormatting(range, color, style, checkpoint: true)) return false;
        this.NotifyMetadataChanged();
        return true;
    }

    internal bool SetExistingBorderFormatting(CellRange range, Color? color, string? style, bool checkpoint)
    {
        if (!this.CanFormatRange(range) || color is null && style is null) return false;
        range = this.ExpandFormattingRange(range);
        if (!this.CanFormatRange(range)) return false;
        Dictionary<CellViewModel, CellAppearance> changes = new();
        CellEdge? Update(CellEdge? edge, BorderSides side, BorderSides sides) => edge is not null && (side & sides) != 0
            ? edge with { Color = color ?? edge.Color, Style = style ?? edge.Style } : edge;
        foreach ((CellViewModel cell, BorderSides sides) in this.BorderTargets(range))
        {
            if (cell.Appearance is not { } original) continue;
            CellAppearance appearance = changes.TryGetValue(cell, out CellAppearance? pending) ? pending : original;
            CellAppearance updated = appearance with
            {
                Left = Update(appearance.Left, BorderSides.Left, sides),
                Right = Update(appearance.Right, BorderSides.Right, sides),
                Top = Update(appearance.Top, BorderSides.Top, sides),
                Bottom = Update(appearance.Bottom, BorderSides.Bottom, sides),
            };
            if (updated != original) changes[cell] = updated;
        }
        if (changes.Count == 0) return false;
        if (checkpoint) this.Checkpoint();
        foreach ((CellViewModel cell, CellAppearance appearance) in changes) cell.Appearance = appearance;
        return true;
    }

    internal void SetBorders(CellRange range, CellBorderPreset preset, CellEdge edge)
    {
        if (!this.CanFormatRange(range)) return;
        range = this.ExpandFormattingRange(range);
        int bottom = range.Row + range.RowSpan - 1;
        int right = range.Column + range.ColumnSpan - 1;
        for (int row = range.Row; row <= bottom; row++)
        {
            for (int column = range.Column; column <= right; column++)
            {
                CellViewModel cell = this.Rows[row].Cells[column];
                CellAppearance appearance = cell.Appearance ?? new();
                bool leftEdge = column == range.Column;
                bool rightEdge = column == right;
                bool topEdge = row == range.Row;
                bool bottomEdge = row == bottom;
                MergeRange? merged = null;
                foreach (MergeRange candidate in this.Merges)
                    if (row >= candidate.Row && row < candidate.Row + candidate.RowSpan
                        && column >= candidate.Column && column < candidate.Column + candidate.ColumnSpan) { merged = candidate; break; }
                bool Set(bool outside, CellBorderPreset side) => preset == CellBorderPreset.None || preset == CellBorderPreset.All
                    || preset == CellBorderPreset.Outside && outside || preset == CellBorderPreset.Inside && !outside || preset == side && outside;
                CellEdge? value = preset == CellBorderPreset.None ? null : edge;
                bool Visible(CellBorderPreset side) => preset == CellBorderPreset.None || merged is not { } merge || side switch
                {
                    CellBorderPreset.Left => column == merge.Column,
                    CellBorderPreset.Right => column == merge.Column + merge.ColumnSpan - 1,
                    CellBorderPreset.Top => row == merge.Row,
                    _ => row == merge.Row + merge.RowSpan - 1,
                };
                cell.Appearance = appearance with
                {
                    Left = Set(leftEdge, CellBorderPreset.Left) && Visible(CellBorderPreset.Left) ? value : appearance.Left,
                    Right = Set(rightEdge, CellBorderPreset.Right) && Visible(CellBorderPreset.Right) ? value : appearance.Right,
                    Top = Set(topEdge, CellBorderPreset.Top) && Visible(CellBorderPreset.Top) ? value : appearance.Top,
                    Bottom = Set(bottomEdge, CellBorderPreset.Bottom) && Visible(CellBorderPreset.Bottom) ? value : appearance.Bottom,
                };
                // A shared edge can also be stored on the adjacent cell, including outside the selection.
                void MatchNeighbour(int neighbourRow, int neighbourColumn, bool horizontal, bool before, bool change)
                {
                    if (!change || !this.IsValid(new CellAddress(neighbourRow, neighbourColumn))) return;
                    CellViewModel neighbour = this.Rows[neighbourRow].Cells[neighbourColumn];
                    if (neighbour.Appearance is not { } other) return;
                    neighbour.Appearance = horizontal
                        ? before ? other with { Right = value } : other with { Left = value }
                        : before ? other with { Bottom = value } : other with { Top = value };
                }
                if (leftEdge) MatchNeighbour(row, column - 1, true, true, Set(true, CellBorderPreset.Left));
                if (rightEdge) MatchNeighbour(row, column + 1, true, false, Set(true, CellBorderPreset.Right));
                if (topEdge) MatchNeighbour(row - 1, column, false, true, Set(true, CellBorderPreset.Top));
                if (bottomEdge) MatchNeighbour(row + 1, column, false, false, Set(true, CellBorderPreset.Bottom));
            }
        }
    }
}
