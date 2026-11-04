namespace SpreadsheetEditor.Controls.ViewModels;

using System;
using ClosedXML.Parser;
using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Models;

public partial class SpreadsheetViewModel
{
    public bool MatchesRule(SheetCellValueRule rule, int row, int column)
    {
        CellViewModel? cell = this.GetCell(new CellAddress(row, column));
        if (cell is null) return false;
        if (rule.Expression is { } expression)
        {
            try
            {
                string relative = FormulaConverter.ToR1C1(expression, rule.Range.Row + 1, rule.Range.Column + 1);
                string adjusted = FormulaConverter.ToA1(relative, row + 1, column + 1);
                FormulaEvaluator evaluator = new((sheet, address) =>
                {
                    CellViewModel? referenced = this.GetCell(address);
                    if (referenced is null) return CellValue.Empty;
                    if (referenced.CachedFormulaValue is { } cached) return cached;
                    if (referenced.NumericValue is { } n) return CellValue.FromNumber(n);
                    return CellValue.FromText(referenced.DisplayText);
                });
                CellValue result = evaluator.Eval(FormulaParser.Parse(adjusted));
                return result.Kind == CellValueKind.Boolean ? result.Boolean : result.TryAsNumber(out double n) && n != 0;
            }
            catch (Exception e) when (e is FormulaParseException or FormulaRuntimeException or ParsingException) { return false; }
        }
        if (rule.TextThreshold is { } text)
        {
            bool equal = cell.DisplayText.Equals(text, StringComparison.OrdinalIgnoreCase);
            return rule.Comparison == CellValueComparison.NotEqual ? !equal : equal;
        }
        return cell.NumericValue is { } number && (rule.IconSet is not null || rule.Matches(number));
    }
}
