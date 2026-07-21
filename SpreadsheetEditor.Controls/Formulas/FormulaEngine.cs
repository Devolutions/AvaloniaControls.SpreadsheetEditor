namespace SpreadsheetEditor.Controls.Formulas;

using System;
using System.Collections.Generic;
using System.Globalization;
using global::SpreadsheetEditor.Controls.Models;

// Computes cell values for a sheet. Reads each cell's raw text through a callback; a cell whose text
// starts with '=' is a formula, anything else is a literal (number / TRUE-FALSE / text). Evaluation
// is recursive with per-pass memoization and cycle detection, so dependencies resolve on demand.
// v1 recomputes everything on each change (simple + correct for these sheet sizes).
public sealed class FormulaEngine
{
    private readonly Func<CellAddress, string> readRaw;

    // raw body -> FormulaNode or FormulaParseException
    private readonly Dictionary<string, object> astCache = new();
    private readonly Dictionary<CellAddress, CellValue> memo = new();
    private readonly HashSet<CellAddress> visiting = new();
    private readonly FormulaEvaluator evaluator;

    private readonly Func<string, string?>? resolveNameFormula;

    public FormulaEngine(Func<CellAddress, string> readRaw, Func<string, string?>? resolveName = null)
    {
        this.readRaw = readRaw;
        this.resolveNameFormula = resolveName;

        // refs resolve back through this engine
        this.evaluator = new FormulaEvaluator(this.Evaluate, this.ResolveNameNode);
    }

    // True if a raw cell string is a formula (starts with '=' and has a body).
    public static bool IsFormula(string? raw) => raw is { Length: > 1 } && raw[0] == '=';

    // Recomputes the given cells; calls apply with the computed value for
    // formula cells and null for non-formula cells (so the caller can clear stale results).
    public void Recalculate(IEnumerable<CellAddress> cells, Action<CellAddress, CellValue?> apply)
    {
        this.memo.Clear();
        this.visiting.Clear();
        foreach (CellAddress address in cells)
        {
            string raw = this.readRaw(address);
            apply(address, IsFormula(raw) ? this.Evaluate(address) : null);
        }
    }

    // Resolves a cell to its value (formula result or parsed literal), memoized per pass.
    public CellValue Evaluate(CellAddress address)
    {
        if (this.memo.TryGetValue(address, out CellValue cached))
        {
            return cached;
        }

        if (this.visiting.Contains(address))
        {
            // cycle — don't memoize the transient
            return CellValue.FromError(FormulaErrors.Circular);
        }

        this.visiting.Add(address);
        CellValue result;
        try
        {
            string raw = this.readRaw(address);
            if (IsFormula(raw))
            {
                object ast = this.GetAst(raw[1..]);
                result = ast is FormulaParseException parseError
                    ? CellValue.FromError(parseError.Error)
                    : this.evaluator.Eval((FormulaNode)ast);
            }
            else
            {
                result = LiteralValue(raw);
            }
        }
        catch (FormulaRuntimeException e)
        {
            result = CellValue.FromError(e.Error);
        }
        catch (FormulaParseException e)
        {
            result = CellValue.FromError(e.Error);
        }
        finally
        {
            this.visiting.Remove(address);
        }

        this.memo[address] = result;
        return result;
    }

    // Parses a defined name's formula into an AST (cached), or null if unknown/unparseable.
    private FormulaNode? ResolveNameNode(string name)
    {
        string? formula = this.resolveNameFormula?.Invoke(name);
        if (string.IsNullOrWhiteSpace(formula))
        {
            return null;
        }

        // FormulaParseException → null
        return this.GetAst(formula) as FormulaNode;
    }

    private object GetAst(string body)
    {
        if (this.astCache.TryGetValue(body, out object? cached))
        {
            return cached;
        }

        object parsed;
        try
        {
            parsed = FormulaParser.Parse(body);
        }
        catch (FormulaParseException e)
        {
            parsed = e;
        }

        this.astCache[body] = parsed;
        return parsed;
    }

    private static CellValue LiteralValue(string raw)
    {
        if (string.IsNullOrEmpty(raw))
        {
            return CellValue.Empty;
        }

        if (double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double n))
        {
            return CellValue.FromNumber(n);
        }

        if (raw.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
        {
            return CellValue.FromBool(true);
        }

        if (raw.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
        {
            return CellValue.FromBool(false);
        }

        return CellValue.FromText(raw);
    }
}
