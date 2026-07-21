namespace SpreadsheetEditor.Controls.Formulas;

using System;
using System.Collections.Generic;
using System.Globalization;
using global::SpreadsheetEditor.Controls.Models;

// Thrown during evaluation to short-circuit with an Excel error literal; caught by the engine.
internal sealed class FormulaRuntimeException(string error) : Exception(error)
{
    public string Error { get; } = error;
}

// Walks a parsed FormulaNode to a CellValue. Cell references are resolved through a callback (so the
// engine can memoize and detect cycles). Errors propagate by throwing FormulaRuntimeException, which
// lets IFERROR catch them naturally.
public sealed class FormulaEvaluator(Func<CellAddress, CellValue> resolveCell, Func<string, FormulaNode?>? resolveName = null)
{
    private readonly HashSet<string> namesVisiting = new(StringComparer.OrdinalIgnoreCase);

    public CellValue Eval(FormulaNode node)
    {
        switch (node)
        {
            case NumberNode n:
                return CellValue.FromNumber(n.Value);
            case TextNode t:
                return CellValue.FromText(t.Value);
            case BoolNode b:
                return CellValue.FromBool(b.Value);

            case NameNode name:
                return this.EvalName(name.Name);

            case CellNode c:
                CellValue v = resolveCell(c.Address);
                if (v.IsError)
                {
                    throw new FormulaRuntimeException(v.Text);
                }

                return v;

            case RangeNode:
                // a range can't collapse to a scalar
                throw new FormulaRuntimeException(FormulaErrors.Value);

            case PercentNode p:
                return CellValue.FromNumber(this.Number(p.Operand) / 100.0);

            case UnaryNode u:
                double operand = this.Number(u.Operand);
                return CellValue.FromNumber(u.Op == "-" ? -operand : operand);

            case BinaryNode bin:
                return this.EvalBinary(bin);

            case FunctionNode f:
                return FormulaFunctions.Call(this, f.Name, f.Args);

            default:
                throw new FormulaRuntimeException(FormulaErrors.Value);
        }
    }

    // ----- Helpers used by operators and functions -----

    // Evaluates a node and coerces to a number, throwing #VALUE! on non-numeric text.
    public double Number(FormulaNode node)
    {
        CellValue v = this.Eval(node);
        if (!v.TryAsNumber(out double n))
        {
            throw new FormulaRuntimeException(FormulaErrors.Value);
        }

        return n;
    }

    public string Text(FormulaNode node) => this.Eval(node).AsText();

    public bool Bool(FormulaNode node)
    {
        CellValue v = this.Eval(node);
        if (v.Kind == CellValueKind.Boolean)
        {
            return v.Boolean;
        }

        if (v.TryAsNumber(out double n))
        {
            return n != 0;
        }

        throw new FormulaRuntimeException(FormulaErrors.Value);
    }

    // Yields every value an argument contributes — one for a scalar, the cells of a range.
    // Range cells are resolved directly (errors surface as error values for the caller to handle).
    public IEnumerable<CellValue> Flatten(FormulaNode node)
    {
        // A name used in a range context (e.g. SUM(MyRange)) expands to its definition's cells.
        if (node is NameNode nm)
        {
            foreach (CellValue v in this.Flatten(this.ResolveName(nm.Name)))
            {
                yield return v;
            }

            yield break;
        }

        if (node is RangeNode range)
        {
            int r1 = Math.Min(range.From.Row, range.To.Row);
            int r2 = Math.Max(range.From.Row, range.To.Row);
            int c1 = Math.Min(range.From.Column, range.To.Column);
            int c2 = Math.Max(range.From.Column, range.To.Column);
            for (int r = r1; r <= r2; r++)
            {
                for (int c = c1; c <= c2; c++)
                {
                    yield return resolveCell(new CellAddress(r, c));
                }
            }
        }
        else
        {
            yield return this.Eval(node);
        }
    }

    // Collects numeric values from args (for SUM/AVERAGE/etc.); errors propagate, blanks and
    // non-numeric text are ignored, booleans count as 1/0.
    public List<double> CollectNumbers(IReadOnlyList<FormulaNode> args)
    {
        List<double> numbers = new();
        foreach (FormulaNode arg in args)
        {
            foreach (CellValue value in this.Flatten(arg))
            {
                if (value.IsError)
                {
                    throw new FormulaRuntimeException(value.Text);
                }

                switch (value.Kind)
                {
                    case CellValueKind.Number:
                        numbers.Add(value.Number);
                        break;
                    case CellValueKind.Boolean:
                        numbers.Add(value.Boolean ? 1 : 0);
                        break;
                    case CellValueKind.Text:
                        if (double.TryParse(value.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out double t))
                        {
                            numbers.Add(t);
                        }

                        break;
                }
            }
        }

        return numbers;
    }

    // Resolves a defined name to its formula AST (#NAME? if unknown).
    private FormulaNode ResolveName(string name) =>
        resolveName?.Invoke(name) ?? throw new FormulaRuntimeException(FormulaErrors.Name);

    // Evaluates a defined name's formula, guarding against name→name cycles.
    private CellValue EvalName(string name)
    {
        if (!this.namesVisiting.Add(name))
        {
            throw new FormulaRuntimeException(FormulaErrors.Circular);
        }

        try
        {
            return this.Eval(this.ResolveName(name));
        }
        finally
        {
            this.namesVisiting.Remove(name);
        }
    }

    private CellValue EvalBinary(BinaryNode bin)
    {
        switch (bin.Op)
        {
            case "&":
                return CellValue.FromText(this.Text(bin.Left) + this.Text(bin.Right));
            case "+":
                return CellValue.FromNumber(this.Number(bin.Left) + this.Number(bin.Right));
            case "-":
                return CellValue.FromNumber(this.Number(bin.Left) - this.Number(bin.Right));
            case "*":
                return CellValue.FromNumber(this.Number(bin.Left) * this.Number(bin.Right));
            case "/":
                double denom = this.Number(bin.Right);
                if (denom == 0)
                {
                    throw new FormulaRuntimeException(FormulaErrors.Div0);
                }

                return CellValue.FromNumber(this.Number(bin.Left) / denom);
            case "^":
                double result = Math.Pow(this.Number(bin.Left), this.Number(bin.Right));
                if (double.IsNaN(result))
                {
                    throw new FormulaRuntimeException(FormulaErrors.Num);
                }

                return CellValue.FromNumber(result);
            default:
                return Compare(bin.Op, this.Eval(bin.Left), this.Eval(bin.Right));
        }
    }

    private static CellValue Compare(string op, CellValue a, CellValue b)
    {
        int cmp = CompareValues(a, b);
        return CellValue.FromBool(op switch
        {
            "=" => cmp == 0,
            "<>" => cmp != 0,
            "<" => cmp < 0,
            ">" => cmp > 0,
            "<=" => cmp <= 0,
            ">=" => cmp >= 0,
            _ => throw new FormulaRuntimeException(FormulaErrors.Value),
        });
    }

    // Excel ordering: numbers < text < booleans; within text, case-insensitive.
    private static int CompareValues(CellValue a, CellValue b)
    {
        (int ra, double na, string ta, bool ba) = Classify(a);
        (int rb, double nb, string tb, bool bb) = Classify(b);
        if (ra != rb)
        {
            return ra.CompareTo(rb);
        }

        return ra switch
        {
            0 => na.CompareTo(nb),
            1 => string.Compare(ta, tb, StringComparison.OrdinalIgnoreCase),
            _ => ba.CompareTo(bb),
        };
    }

    private static (int Rank, double Num, string Text, bool Boolean) Classify(CellValue v) => v.Kind switch
    {
        CellValueKind.Boolean => (2, 0, string.Empty, v.Boolean),
        CellValueKind.Text => (1, 0, v.Text, false),
        CellValueKind.Empty => (0, 0, string.Empty, false),
        _ => (0, v.Number, string.Empty, false),
    };
}
