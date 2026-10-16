namespace SpreadsheetEditor.Controls.Formulas;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

// The built-in worksheet functions. Each takes the evaluator and the (unevaluated) arg
// nodes, so functions can flatten ranges or evaluate lazily (IF/IFERROR).
public static class FormulaFunctions
{
    private delegate CellValue Fn(FormulaEvaluator ev, IReadOnlyList<FormulaNode> args);

    private static readonly Dictionary<string, Fn> Functions = new(StringComparer.OrdinalIgnoreCase)
    {
        ["SUM"] = static (ev, a) => CellValue.FromNumber(ev.CollectNumbers(a).Sum()),
        ["AVERAGE"] = static (ev, a) =>
        {
            List<double> n = ev.CollectNumbers(a);
            return n.Count == 0 ? Err(FormulaErrors.Div0) : CellValue.FromNumber(n.Average());
        },
        ["MIN"] = static (ev, a) =>
        {
            List<double> n = ev.CollectNumbers(a);
            return CellValue.FromNumber(n.Count == 0 ? 0 : n.Min());
        },
        ["MAX"] = static (ev, a) =>
        {
            List<double> n = ev.CollectNumbers(a);
            return CellValue.FromNumber(n.Count == 0 ? 0 : n.Max());
        },
        ["COUNT"] = static (ev, a) => CellValue.FromNumber(CountWhere(ev, a, static v => v.Kind == CellValueKind.Number)),
        ["COUNTA"] = static (ev, a) => CellValue.FromNumber(CountWhere(ev, a, static v => v.Kind != CellValueKind.Empty)),
        ["COUNTIF"] = static (ev, a) =>
        {
            Arity(a, 2, 2);
            CellValue criterion = ev.Eval(a[1]);
            return CellValue.FromNumber(ev.Flatten(a[0]).Count(value => MatchesCriterion(value, criterion)));
        },

        // SUBTOTAL(function_num, ref, ...). function_num 1-11 (and 101-111, which ignore manually
        // hidden rows — treated identically here since the grid has no hidden rows) selects the
        // aggregate. Used by Excel table Totals rows.
        ["SUBTOTAL"] = static (ev, a) =>
        {
            if (a.Count < 2)
            {
                throw Throw(FormulaErrors.Value);
            }

            int function = (int)Math.Round(ev.Number(a[0])) % 100;
            List<FormulaNode> rest = a.Skip(1).ToList();
            List<double> n = ev.CollectNumbers(rest);
            return function switch
            {
                1 => n.Count == 0 ? Err(FormulaErrors.Div0) : CellValue.FromNumber(n.Average()),        // AVERAGE
                2 => CellValue.FromNumber(CountWhere(ev, rest, static v => v.Kind == CellValueKind.Number)), // COUNT
                3 => CellValue.FromNumber(CountWhere(ev, rest, static v => v.Kind != CellValueKind.Empty)),  // COUNTA
                4 => CellValue.FromNumber(n.Count == 0 ? 0 : n.Max()),                                  // MAX
                5 => CellValue.FromNumber(n.Count == 0 ? 0 : n.Min()),                                  // MIN
                6 => CellValue.FromNumber(n.Aggregate(1.0, static (x, y) => x * y)),                    // PRODUCT
                9 => CellValue.FromNumber(n.Sum()),                                                     // SUM
                _ => throw Throw(FormulaErrors.Value),                                                  // STDEV/VAR: unsupported
            };
        },

        ["IF"] = static (ev, a) =>
        {
            Arity(a, 2, 3);
            if (ev.Bool(a[0]))
            {
                return ev.Eval(a[1]);
            }

            return a.Count == 3 ? ev.Eval(a[2]) : CellValue.FromBool(false);
        },
        ["IFERROR"] = static (ev, a) =>
        {
            Arity(a, 2, 2);
            try
            {
                return ev.Eval(a[0]);
            }
            catch (FormulaRuntimeException)
            {
                return ev.Eval(a[1]);
            }
        },
        ["AND"] = static (ev, a) => CellValue.FromBool(BoolsFrom(ev, a).All(static b => b)),
        ["OR"] = static (ev, a) => CellValue.FromBool(BoolsFrom(ev, a).Any(static b => b)),
        ["NOT"] = static (ev, a) =>
        {
            Arity(a, 1, 1);
            return CellValue.FromBool(!ev.Bool(a[0]));
        },

        ["ABS"] = static (ev, a) => One(a, _ => Math.Abs(ev.Number(a[0]))),
        ["SQRT"] = static (ev, a) => One(a, _ =>
        {
            double x = ev.Number(a[0]);
            if (x < 0)
            {
                throw Throw(FormulaErrors.Num);
            }

            return Math.Sqrt(x);
        }),
        ["INT"] = static (ev, a) => One(a, _ => Math.Floor(ev.Number(a[0]))),
        ["MOD"] = static (ev, a) =>
        {
            Arity(a, 2, 2);
            double d = ev.Number(a[1]);
            if (d == 0)
            {
                throw Throw(FormulaErrors.Div0);
            }

            double n = ev.Number(a[0]);

            // Excel MOD takes the divisor's sign
            return CellValue.FromNumber(n - (Math.Floor(n / d) * d));
        },
        ["POWER"] = static (ev, a) =>
        {
            Arity(a, 2, 2);
            double r = Math.Pow(ev.Number(a[0]), ev.Number(a[1]));
            if (double.IsNaN(r))
            {
                throw Throw(FormulaErrors.Num);
            }

            return CellValue.FromNumber(r);
        },
        ["ROUND"] = static (ev, a) => Round(ev, a, MidpointRounding.AwayFromZero),
        ["ROUNDUP"] = static (ev, a) =>
        {
            Arity(a, 2, 2);
            double f = Math.Pow(10, ev.Number(a[1]));
            double x = ev.Number(a[0]);
            return CellValue.FromNumber(Math.Sign(x) * Math.Ceiling(Math.Abs(x) * f) / f);
        },
        ["ROUNDDOWN"] = static (ev, a) =>
        {
            Arity(a, 2, 2);
            double f = Math.Pow(10, ev.Number(a[1]));
            double x = ev.Number(a[0]);
            return CellValue.FromNumber(Math.Sign(x) * Math.Floor(Math.Abs(x) * f) / f);
        },

        // Dates: Excel stores dates as serial numbers (OLE Automation date, epoch 1899-12-30).
        ["TODAY"] = static (ev, a) =>
        {
            Arity(a, 0, 0);
            return CellValue.FromNumber(Math.Floor(DateTime.Today.ToOADate()));
        },
        ["NOW"] = static (ev, a) =>
        {
            Arity(a, 0, 0);
            return CellValue.FromNumber(DateTime.Now.ToOADate());
        },
        ["YEAR"] = static (ev, a) => DatePart(ev, a, static d => d.Year),
        ["MONTH"] = static (ev, a) => DatePart(ev, a, static d => d.Month),
        ["DAY"] = static (ev, a) => DatePart(ev, a, static d => d.Day),
        ["DATE"] = static (ev, a) =>
        {
            Arity(a, 3, 3);
            try
            {
                return CellValue.FromNumber(new DateTime((int)ev.Number(a[0]), (int)ev.Number(a[1]), (int)ev.Number(a[2])).ToOADate());
            }
            catch
            {
                throw Throw(FormulaErrors.Num);
            }
        },

        ["CONCAT"] = static (ev, a) => Concat(ev, a),
        ["CONCATENATE"] = static (ev, a) => Concat(ev, a),
        ["LEN"] = static (ev, a) =>
        {
            Arity(a, 1, 1);
            return CellValue.FromNumber(ev.Text(a[0]).Length);
        },
        ["UPPER"] = static (ev, a) =>
        {
            Arity(a, 1, 1);
            return CellValue.FromText(ev.Text(a[0]).ToUpperInvariant());
        },
        ["LOWER"] = static (ev, a) =>
        {
            Arity(a, 1, 1);
            return CellValue.FromText(ev.Text(a[0]).ToLowerInvariant());
        },
        ["TRIM"] = static (ev, a) =>
        {
            Arity(a, 1, 1);
            return CellValue.FromText(ev.Text(a[0]).Trim());
        },
        ["LEFT"] = static (ev, a) =>
        {
            Arity(a, 1, 2);
            string s = ev.Text(a[0]);
            int n = a.Count == 2 ? (int)ev.Number(a[1]) : 1;
            return CellValue.FromText(s[..Math.Clamp(n, 0, s.Length)]);
        },
        ["RIGHT"] = static (ev, a) =>
        {
            Arity(a, 1, 2);
            string s = ev.Text(a[0]);
            int n = a.Count == 2 ? (int)ev.Number(a[1]) : 1;
            n = Math.Clamp(n, 0, s.Length);
            return CellValue.FromText(s[(s.Length - n)..]);
        },
        ["MID"] = static (ev, a) =>
        {
            Arity(a, 3, 3);
            string s = ev.Text(a[0]);
            int start = (int)ev.Number(a[1]);
            int len = (int)ev.Number(a[2]);
            if (start < 1 || len < 0)
            {
                throw Throw(FormulaErrors.Value);
            }

            if (start > s.Length)
            {
                return CellValue.FromText(string.Empty);
            }

            return CellValue.FromText(s.Substring(start - 1, Math.Min(len, s.Length - start + 1)));
        },
    };

    public static CellValue Call(FormulaEvaluator ev, string name, IReadOnlyList<FormulaNode> args)
    {
        if (!Functions.TryGetValue(name, out Fn? fn))
        {
            throw Throw(FormulaErrors.Name);
        }

        return fn(ev, args);
    }

    // ----- shared helpers -----

    private static CellValue DatePart(FormulaEvaluator ev, IReadOnlyList<FormulaNode> a, Func<DateTime, int> part)
    {
        Arity(a, 1, 1);
        CellValue value = ev.Eval(a[0]);
        // Imported date cells retain their formatted text. Parse those directly rather than
        // throwing a numeric-conversion exception for every row in a large date-helper column.
        if (value.Kind == CellValueKind.Text &&
            DateTime.TryParse(value.Text, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out DateTime date))
        {
            return CellValue.FromNumber(part(date));
        }

        if (!value.TryAsNumber(out double serial) || !double.IsFinite(serial) ||
            serial <= -657435 || serial >= 2958466)
        {
            throw Throw(FormulaErrors.Num);
        }
        return CellValue.FromNumber(part(DateTime.FromOADate(serial)));
    }

    private static CellValue Round(FormulaEvaluator ev, IReadOnlyList<FormulaNode> a, MidpointRounding mode)
    {
        Arity(a, 2, 2);
        int digits = (int)ev.Number(a[1]);
        double x = ev.Number(a[0]);
        double f = Math.Pow(10, digits);
        return CellValue.FromNumber(Math.Round(x * f, mode) / f);
    }

    private static CellValue Concat(FormulaEvaluator ev, IReadOnlyList<FormulaNode> args)
    {
        System.Text.StringBuilder sb = new();
        foreach (FormulaNode arg in args)
        {
            foreach (CellValue v in ev.Flatten(arg))
            {
                if (v.IsError)
                {
                    throw Throw(v.Text);
                }

                sb.Append(v.AsText());
            }
        }

        return CellValue.FromText(sb.ToString());
    }

    private static IEnumerable<bool> BoolsFrom(FormulaEvaluator ev, IReadOnlyList<FormulaNode> args)
    {
        foreach (FormulaNode arg in args)
        {
            foreach (CellValue v in ev.Flatten(arg))
            {
                if (v.IsError)
                {
                    throw Throw(v.Text);
                }

                if (v.Kind == CellValueKind.Empty)
                {
                    continue;
                }

                if (v.Kind == CellValueKind.Boolean)
                {
                    yield return v.Boolean;
                    continue;
                }

                if (v.TryAsNumber(out double n))
                {
                    yield return n != 0;
                }
            }
        }
    }

    private static int CountWhere(FormulaEvaluator ev, IReadOnlyList<FormulaNode> args, Func<CellValue, bool> predicate)
    {
        int count = 0;
        foreach (FormulaNode arg in args)
        {
            foreach (CellValue v in ev.Flatten(arg))
            {
                if (v.IsError)
                {
                    throw Throw(v.Text);
                }

                if (predicate(v))
                {
                    count++;
                }
            }
        }

        return count;
    }

    private static bool MatchesCriterion(CellValue value, CellValue criterion)
    {
        if (criterion.Kind == CellValueKind.Boolean)
        {
            return value.Kind == CellValueKind.Boolean && value.Boolean == criterion.Boolean;
        }

        if (criterion.Kind == CellValueKind.Number)
        {
            return value.TryAsNumber(out double number) && number == criterion.Number;
        }

        string expression = criterion.AsText();
        string op = expression.StartsWith("<=", StringComparison.Ordinal) ||
            expression.StartsWith(">=", StringComparison.Ordinal) ||
            expression.StartsWith("<>", StringComparison.Ordinal)
                ? expression[..2]
                : expression.StartsWith('<') || expression.StartsWith('>') || expression.StartsWith('=')
                    ? expression[..1]
                    : "=";
        string expected = op == "=" && !expression.StartsWith('=') ? expression : expression[op.Length..];

        int comparison;
        if (double.TryParse(expected, NumberStyles.Any, CultureInfo.InvariantCulture, out double expectedNumber) &&
            value.TryAsNumber(out double actualNumber))
        {
            comparison = actualNumber.CompareTo(expectedNumber);
        }
        else if (bool.TryParse(expected, out bool expectedBoolean) && value.Kind == CellValueKind.Boolean)
        {
            comparison = value.Boolean.CompareTo(expectedBoolean);
        }
        else
        {
            comparison = string.Compare(value.AsText(), expected, StringComparison.OrdinalIgnoreCase);
        }

        return op switch
        {
            "=" => comparison == 0,
            "<>" => comparison != 0,
            "<" => comparison < 0,
            ">" => comparison > 0,
            "<=" => comparison <= 0,
            ">=" => comparison >= 0,
            _ => false,
        };
    }

    private static CellValue One(IReadOnlyList<FormulaNode> a, Func<FormulaNode, double> compute)
    {
        Arity(a, 1, 1);
        return CellValue.FromNumber(compute(a[0]));
    }

    private static void Arity(IReadOnlyList<FormulaNode> args, int min, int max)
    {
        if (args.Count < min || args.Count > max)
        {
            throw Throw(FormulaErrors.Value);
        }
    }

    private static CellValue Err(string error) => CellValue.FromError(error);

    private static FormulaRuntimeException Throw(string error) => new(error);
}
