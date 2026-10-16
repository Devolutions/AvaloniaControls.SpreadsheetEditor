namespace SpreadsheetEditor.Controls.Formulas;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

using global::SpreadsheetEditor.Controls.Models;

// Thrown on a malformed formula; carries the Excel error literal to surface.
public sealed class FormulaParseException(string error) : Exception(error)
{
    public string Error { get; } = error;
}

// Tokenizes and parses a formula (the text after the leading '=') into a FormulaNode.
// Recursive-descent with Excel precedence: comparison < concat < +/- < *// < ^ < unary
// < postfix % < primary. (Unary binds tighter than ^, so -2^2 = 4, matching Excel.)
public static class FormulaParser
{
    private static readonly Regex CellRefPattern =
        new(@"^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}$", RegexOptions.Compiled);

    public static FormulaNode Parse(string formulaBody)
    {
        List<Token> tokens = Tokenize(formulaBody);
        Cursor parser = new(tokens);
        FormulaNode node = parser.ParseComparison();
        if (!parser.AtEnd)
        {
            throw new FormulaParseException(FormulaErrors.Value);
        }

        return node;
    }

    // Returns worksheet names referenced by a formula. The lazy workbook loader uses this to load
    // formula dependencies even when their tabs have not been selected yet.
    public static IReadOnlySet<string> ReferencedSheets(string formulaBody)
    {
        HashSet<string> sheets = new(StringComparer.OrdinalIgnoreCase);
        CollectReferencedSheets(Parse(formulaBody), sheets);
        return sheets;
    }

    internal static IEnumerable<(string? Sheet, CellRange Range)> ReferencedRanges(
        string formulaBody, Func<string, string?> resolveName)
    {
        List<(string?, CellRange)> references = new();
        HashSet<string> visiting = new(StringComparer.OrdinalIgnoreCase);
        void Collect(FormulaNode node)
        {
            switch (node)
            {
                case CellNode cell:
                    references.Add((cell.SheetName, new CellRange(cell.Address.Row, cell.Address.Column, 1, 1)));
                    break;
                case RangeNode range:
                    references.Add((range.SheetName, new CellRange(
                        Math.Min(range.From.Row, range.To.Row), Math.Min(range.From.Column, range.To.Column),
                        Math.Abs(range.To.Row - range.From.Row) + 1, Math.Abs(range.To.Column - range.From.Column) + 1)));
                    break;
                case NameNode name when visiting.Add(name.Name):
                    if (resolveName(name.Name) is { Length: > 0 } definition)
                    {
                        Collect(Parse(definition.TrimStart('=')));
                    }
                    visiting.Remove(name.Name);
                    break;
                case UnaryNode unary:
                    Collect(unary.Operand);
                    break;
                case PercentNode percent:
                    Collect(percent.Operand);
                    break;
                case BinaryNode binary:
                    Collect(binary.Left);
                    Collect(binary.Right);
                    break;
                case FunctionNode function:
                    if (function.Name is "TODAY" or "NOW")
                    {
                        // '*' cannot be an Excel worksheet name: reserve it for volatile inputs.
                        references.Add(("*", new CellRange(0, 0, int.MaxValue, int.MaxValue)));
                    }
                    foreach (FormulaNode argument in function.Args)
                    {
                        Collect(argument);
                    }
                    break;
            }
        }
        Collect(Parse(formulaBody));
        return references;
    }

    private static void CollectReferencedSheets(FormulaNode node, ISet<string> sheets)
    {
        switch (node)
        {
            case CellNode { SheetName: { Length: > 0 } sheetName }:
                sheets.Add(sheetName);
                break;
            case RangeNode { SheetName: { Length: > 0 } sheetName }:
                sheets.Add(sheetName);
                break;
            case UnaryNode unary:
                CollectReferencedSheets(unary.Operand, sheets);
                break;
            case PercentNode percent:
                CollectReferencedSheets(percent.Operand, sheets);
                break;
            case BinaryNode binary:
                CollectReferencedSheets(binary.Left, sheets);
                CollectReferencedSheets(binary.Right, sheets);
                break;
            case FunctionNode function:
                foreach (FormulaNode argument in function.Args)
                {
                    CollectReferencedSheets(argument, sheets);
                }
                break;
        }
    }

    // ----- Tokens -----

    private enum TokKind
    {
        Number,
        Text,
        Word,
        Sheet,
        Bang,
        Op,
        LParen,
        RParen,
        Comma,
        Colon,
        End,
    }

    private static List<Token> Tokenize(string s)
    {
        List<Token> tokens = new();
        int i = 0;
        while (i < s.Length)
        {
            char c = s[i];
            if (char.IsWhiteSpace(c))
            {
                i++;
                continue;
            }

            if (c == '"')
            {
                StringBuilder sb = new();
                i++;
                while (i < s.Length)
                {
                    if (s[i] == '"')
                    {
                        if (i + 1 < s.Length && s[i + 1] == '"')
                        {
                            sb.Append('"');
                            i += 2;
                            continue;
                        }

                        i++;
                        break;
                    }

                    sb.Append(s[i++]);
                }

                tokens.Add(new Token(TokKind.Text, sb.ToString()));
                continue;
            }

            if (char.IsDigit(c) || (c == '.' && i + 1 < s.Length && char.IsDigit(s[i + 1])))
            {
                int start = i;
                while (i < s.Length && (char.IsDigit(s[i]) || s[i] == '.'))
                {
                    i++;
                }

                if (i < s.Length && (s[i] == 'e' || s[i] == 'E'))
                {
                    i++;
                    if (i < s.Length && (s[i] == '+' || s[i] == '-'))
                    {
                        i++;
                    }

                    while (i < s.Length && char.IsDigit(s[i]))
                    {
                        i++;
                    }
                }

                string text = s[start..i];
                if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double num))
                {
                    throw new FormulaParseException(FormulaErrors.Value);
                }

                tokens.Add(new Token(TokKind.Number, text, num));
                continue;
            }

            // A single-quoted sheet name, e.g. 'Profit and Loss' (with '' as an escaped quote).
            if (c == '\'')
            {
                StringBuilder sb = new();
                i++;
                while (i < s.Length)
                {
                    if (s[i] == '\'')
                    {
                        if (i + 1 < s.Length && s[i + 1] == '\'')
                        {
                            sb.Append('\'');
                            i += 2;
                            continue;
                        }

                        i++;
                        break;
                    }

                    sb.Append(s[i++]);
                }

                tokens.Add(new Token(TokKind.Sheet, sb.ToString()));
                continue;
            }

            if (c == '!')
            {
                tokens.Add(new Token(TokKind.Bang, "!"));
                i++;
                continue;
            }

            // A "word": cell ref, function name, TRUE/FALSE (letters, digits, $, .).
            if (char.IsLetter(c) || c == '$' || c == '_')
            {
                int start = i;
                while (i < s.Length && (char.IsLetterOrDigit(s[i]) || s[i] == '$' || s[i] == '_' || s[i] == '.'))
                {
                    i++;
                }

                tokens.Add(new Token(TokKind.Word, s[start..i]));
                continue;
            }

            switch (c)
            {
                case '(':
                    tokens.Add(new Token(TokKind.LParen, "("));
                    i++;
                    continue;
                case ')':
                    tokens.Add(new Token(TokKind.RParen, ")"));
                    i++;
                    continue;
                case ',':
                    tokens.Add(new Token(TokKind.Comma, ","));
                    i++;
                    continue;
                case ':':
                    tokens.Add(new Token(TokKind.Colon, ":"));
                    i++;
                    continue;
            }

            // Operators (multi-char first).
            if (c is '<' or '>')
            {
                if (i + 1 < s.Length && s[i + 1] == '=')
                {
                    tokens.Add(new Token(TokKind.Op, c + "="));
                    i += 2;
                    continue;
                }

                if (c == '<' && i + 1 < s.Length && s[i + 1] == '>')
                {
                    tokens.Add(new Token(TokKind.Op, "<>"));
                    i += 2;
                    continue;
                }
            }

            if ("+-*/^%&=<>".IndexOf(c) >= 0)
            {
                tokens.Add(new Token(TokKind.Op, c.ToString()));
                i++;
                continue;
            }

            throw new FormulaParseException(FormulaErrors.Value);
        }

        tokens.Add(new Token(TokKind.End, string.Empty));
        return tokens;
    }

    // Parses an A1-style reference (with optional $ anchors) to a zero-based address.
    private static bool TryParseCell(string word, out CellAddress address)
    {
        address = default;
        if (!CellRefPattern.IsMatch(word))
        {
            return false;
        }

        string cleaned = word.Replace("$", string.Empty);
        int split = 0;
        while (split < cleaned.Length && char.IsLetter(cleaned[split]))
        {
            split++;
        }

        if (split == 0 || split == cleaned.Length)
        {
            return false;
        }

        string letters = cleaned[..split];
        if (!int.TryParse(cleaned[split..], NumberStyles.None, CultureInfo.InvariantCulture, out int rowNumber) || rowNumber < 1)
        {
            return false;
        }

        try
        {
            address = new CellAddress(rowNumber - 1, CellAddress.LettersToColumn(letters));
            return true;
        }
        catch
        {
            return false;
        }
    }

    private readonly record struct Token(TokKind Kind, string Text, double Number = 0);

    // ----- Recursive-descent parser over the token list -----

    private sealed class Cursor(List<Token> tokens)
    {
        private int pos;

        public bool AtEnd => this.Peek.Kind == TokKind.End;

        private Token Peek => tokens[this.pos];

        public FormulaNode ParseComparison()
        {
            FormulaNode left = this.ParseConcat();
            while (this.TryOp(out string op, "=", "<>", "<", ">", "<=", ">="))
            {
                left = new BinaryNode(op, left, this.ParseConcat());
            }

            return left;
        }

        private Token Next() => tokens[this.pos++];

        private bool TryOp(out string op, params string[] ops)
        {
            if (this.Peek.Kind == TokKind.Op && Array.IndexOf(ops, this.Peek.Text) >= 0)
            {
                op = this.Next().Text;
                return true;
            }

            op = string.Empty;
            return false;
        }

        private FormulaNode ParseConcat()
        {
            FormulaNode left = this.ParseAdditive();
            while (this.TryOp(out string op, "&"))
            {
                left = new BinaryNode(op, left, this.ParseAdditive());
            }

            return left;
        }

        private FormulaNode ParseAdditive()
        {
            FormulaNode left = this.ParseMultiplicative();
            while (this.TryOp(out string op, "+", "-"))
            {
                left = new BinaryNode(op, left, this.ParseMultiplicative());
            }

            return left;
        }

        private FormulaNode ParseMultiplicative()
        {
            FormulaNode left = this.ParsePower();
            while (this.TryOp(out string op, "*", "/"))
            {
                left = new BinaryNode(op, left, this.ParsePower());
            }

            return left;
        }

        private FormulaNode ParsePower()
        {
            FormulaNode left = this.ParseUnary();
            if (this.TryOp(out _, "^"))
            {
                // right-associative
                return new BinaryNode("^", left, this.ParsePower());
            }

            return left;
        }

        private FormulaNode ParseUnary()
        {
            if (this.TryOp(out string op, "-", "+"))
            {
                return new UnaryNode(op, this.ParseUnary());
            }

            return this.ParsePostfix();
        }

        private FormulaNode ParsePostfix()
        {
            FormulaNode node = this.ParsePrimary();
            while (this.TryOp(out _, "%"))
            {
                node = new PercentNode(node);
            }

            return node;
        }

        private FormulaNode ParsePrimary()
        {
            Token tok = this.Peek;
            switch (tok.Kind)
            {
                case TokKind.Number:
                    this.Next();
                    return new NumberNode(tok.Number);
                case TokKind.Text:
                    this.Next();
                    return new TextNode(tok.Text);
                case TokKind.LParen:
                    this.Next();
                    FormulaNode inner = this.ParseComparison();
                    this.Expect(TokKind.RParen);
                    return inner;
                case TokKind.Word:
                    return this.ParseWord();
                case TokKind.Sheet:
                    string quotedSheetName = this.Next().Text;
                    this.Expect(TokKind.Bang);
                    return this.ParseRefWord(quotedSheetName);
                default:
                    throw new FormulaParseException(FormulaErrors.Value);
            }
        }

        private FormulaNode ParseWord()
        {
            string word = this.Next().Text;

            // Sheet-qualified reference: Sheet!ref.
            if (this.Peek.Kind == TokKind.Bang)
            {
                this.Next();
                return this.ParseRefWord(word);
            }

            // Function call: WORD( args )
            if (this.Peek.Kind == TokKind.LParen)
            {
                this.Next();
                List<FormulaNode> args = new();
                if (this.Peek.Kind != TokKind.RParen)
                {
                    args.Add(this.ParseComparison());
                    while (this.Peek.Kind == TokKind.Comma)
                    {
                        this.Next();
                        args.Add(this.ParseComparison());
                    }
                }

                this.Expect(TokKind.RParen);
                return new FunctionNode(word.ToUpperInvariant(), args);
            }

            // TRUE / FALSE
            if (word.Equals("TRUE", StringComparison.OrdinalIgnoreCase))
            {
                return new BoolNode(true);
            }

            if (word.Equals("FALSE", StringComparison.OrdinalIgnoreCase))
            {
                return new BoolNode(false);
            }

            // Cell reference, optionally the start of a range A1:B2.
            if (TryParseCell(word, out CellAddress from))
            {
                if (this.Peek.Kind == TokKind.Colon)
                {
                    this.Next();
                    if (this.Peek.Kind != TokKind.Word || !TryParseCell(this.Next().Text, out CellAddress to))
                    {
                        throw new FormulaParseException(FormulaErrors.Ref);
                    }

                    return new RangeNode(from, to);
                }

                return new CellNode(from);
            }

            return new NameNode(word);
        }

        // Parses a cell/range reference that follows a sheet qualifier (Sheet!A1 / Sheet!A1:B2).
        private FormulaNode ParseRefWord(string sheetName)
        {
            if (this.Peek.Kind != TokKind.Word || !TryParseCell(this.Next().Text, out CellAddress from))
            {
                throw new FormulaParseException(FormulaErrors.Ref);
            }

            if (this.Peek.Kind == TokKind.Colon)
            {
                this.Next();

                // A range across sheets writes the sheet again: A1:'Sheet'!B2 — tolerate an extra qualifier.
                if (this.Peek.Kind == TokKind.Sheet)
                {
                    this.Next();
                    this.Expect(TokKind.Bang);
                }

                if (this.Peek.Kind != TokKind.Word || !TryParseCell(this.Next().Text, out CellAddress to))
                {
                    throw new FormulaParseException(FormulaErrors.Ref);
                }

                return new RangeNode(from, to, sheetName);
            }

            return new CellNode(from, sheetName);
        }

        private void Expect(TokKind kind)
        {
            if (this.Peek.Kind != kind)
            {
                throw new FormulaParseException(FormulaErrors.Value);
            }

            this.Next();
        }
    }
}
