namespace SpreadsheetEditor.Controls.Formulas;

using System.Collections.Generic;
using global::SpreadsheetEditor.Controls.Models;

// Parsed formula AST. Evaluated by FormulaEvaluator.
public abstract record FormulaNode;

public sealed record NumberNode(double Value) : FormulaNode;
public sealed record TextNode(string Value) : FormulaNode;
public sealed record BoolNode(bool Value) : FormulaNode;

// A single cell reference, e.g. A1 / $A$1.
public sealed record CellNode(CellAddress Address) : FormulaNode;

// A rectangular range, e.g. A1:B3.
public sealed record RangeNode(CellAddress From, CellAddress To) : FormulaNode;

// A bareword that isn't a cell ref or TRUE/FALSE — evaluates to #NAME?.
public sealed record NameNode(string Name) : FormulaNode;

public sealed record UnaryNode(string Op, FormulaNode Operand) : FormulaNode;
public sealed record PercentNode(FormulaNode Operand) : FormulaNode;
public sealed record BinaryNode(string Op, FormulaNode Left, FormulaNode Right) : FormulaNode;
public sealed record FunctionNode(string Name, IReadOnlyList<FormulaNode> Args) : FormulaNode;
