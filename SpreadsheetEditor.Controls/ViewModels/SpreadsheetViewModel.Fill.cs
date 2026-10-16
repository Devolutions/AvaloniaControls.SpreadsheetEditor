namespace SpreadsheetEditor.Controls.ViewModels;

using System.Collections.Generic;
using System.Linq;

using ClosedXML.Parser;

using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Models;

public partial class SpreadsheetViewModel
{
    // Repeat a source block into an extended rectangle. The source is never overwritten, and a
    // completed fill is one undo point and one workbook-recalculation request, not one per cell.
    public bool FillCells(CellRange source, CellRange target)
    {
        if (source.Row < 0 || source.Column < 0 || source.RowSpan <= 0 || source.ColumnSpan <= 0 ||
            !this.IsValid(new CellAddress(source.Row + source.RowSpan - 1, source.Column + source.ColumnSpan - 1)) ||
            target.Row < 0 || target.Column < 0 || target.RowSpan <= 0 || target.ColumnSpan <= 0 ||
            target.Row > source.Row || target.Column > source.Column ||
            (long)target.Row + target.RowSpan < (long)source.Row + source.RowSpan ||
            (long)target.Column + target.ColumnSpan < (long)source.Column + source.ColumnSpan ||
            (long)target.Row + target.RowSpan > 1048576 || (long)target.Column + target.ColumnSpan > 16384 ||
            source == target || this.Merges.Any(merge => Intersects(target, merge)))
        {
            return false;
        }

        CellData[][] data = new CellData[source.RowSpan][];
        string?[][] relativeFormulas = new string?[source.RowSpan][];
        for (int row = 0; row < source.RowSpan; row++)
        {
            data[row] = new CellData[source.ColumnSpan];
            relativeFormulas[row] = new string?[source.ColumnSpan];
            for (int column = 0; column < source.ColumnSpan; column++)
            {
                CellData cell = this.Rows[source.Row + row].Cells[source.Column + column].ToData();
                data[row][column] = cell;
                if (FormulaEngine.IsFormula(cell.Text))
                {
                    try
                    {
                        relativeFormulas[row][column] = FormulaConverter.ToR1C1(
                            cell.Text[1..], source.Row + row + 1, source.Column + column + 1);
                    }
                    catch (ParsingException)
                    {
                        // Preserve a malformed/unsupported formula verbatim, just like copying it.
                    }
                }
            }
        }

        this.Checkpoint();
        this.EnsureSize(target.Row + target.RowSpan, target.Column + target.ColumnSpan);
        List<CellAddress> changed = new();
        bool formulaChanged = false;
        bool wasSuspended = this.suspendRecalc;
        this.suspendRecalc = true;
        try
        {
            for (int row = target.Row; row < target.Row + target.RowSpan; row++)
            {
                for (int column = target.Column; column < target.Column + target.ColumnSpan; column++)
                {
                    if (row >= source.Row && row < source.Row + source.RowSpan &&
                        column >= source.Column && column < source.Column + source.ColumnSpan)
                    {
                        continue;
                    }
                    // Positive modulo also repeats the block correctly when filling up or left.
                    int sourceRow = ((row - source.Row) % source.RowSpan + source.RowSpan) % source.RowSpan;
                    int sourceColumn = ((column - source.Column) % source.ColumnSpan + source.ColumnSpan) % source.ColumnSpan;
                    CellData cell = data[sourceRow][sourceColumn];
                    if (relativeFormulas[sourceRow][sourceColumn] is { } relative)
                    {
                        cell = cell with { Text = "=" + FormulaConverter.ToA1(relative, row + 1, column + 1) };
                    }
                    CellViewModel destination = this.Rows[row].Cells[column];
                    bool formula = FormulaEngine.IsFormula(cell.Text);
                    if (cell.Text != destination.Text || formula)
                    {
                        changed.Add(new CellAddress(row, column));
                        formulaChanged |= formula || destination.IsFormula;
                    }
                    if (formula)
                    {
                        // A source's cached result belongs to its original references, not this cell.
                        cell = cell with { FormulaResult = null, FormulaNumericValue = null, FormulaValue = null };
                    }
                    destination.ApplyData(cell);
                }
            }
        }
        finally
        {
            this.suspendRecalc = wasSuspended;
        }
        this.batchedChangedCells.Clear(); // This operation submits its exact changed-address list below.
        if (changed.Count == 0)
        {
            this.NotifyMetadataChanged();
        }
        else if (this.editRecalculationGate?.Invoke(changed, formulaChanged) ?? true)
        {
            this.Recalculate();
        }
        return true;
    }

    private static bool Intersects(CellRange range, MergeRange merge) =>
        range.Row < merge.Row + merge.RowSpan && range.Row + range.RowSpan > merge.Row &&
        range.Column < merge.Column + merge.ColumnSpan && range.Column + range.ColumnSpan > merge.Column;
}
