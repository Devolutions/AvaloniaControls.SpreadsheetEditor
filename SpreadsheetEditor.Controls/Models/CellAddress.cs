namespace SpreadsheetEditor.Controls.Models;

using System;
using System.Text;

// A zero-based (row, column) cell coordinate with A1-style conversion (e.g. (0,0) == "A1").
public readonly record struct CellAddress(int Row, int Column)
{
    // Converts a zero-based column index to its spreadsheet letter(s): 0 -> A, 25 -> Z, 26 -> AA.
    public static string ColumnToLetters(int column)
    {
        if (column < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(column));
        }

        StringBuilder sb = new();
        column++; // shift to 1-based for the modulo math
        while (column > 0)
        {
            column--;
            sb.Insert(0, (char)('A' + (column % 26)));
            column /= 26;
        }

        return sb.ToString();
    }

    // Parses spreadsheet column letter(s) back to a zero-based index: "A" -> 0, "AA" -> 26.
    public static int LettersToColumn(string letters)
    {
        if (string.IsNullOrEmpty(letters))
        {
            throw new ArgumentException("Column letters must not be empty.", nameof(letters));
        }

        int column = 0;
        foreach (char c in letters)
        {
            char upper = char.ToUpperInvariant(c);
            if (upper is < 'A' or > 'Z')
            {
                throw new FormatException($"Invalid column letter '{c}'.");
            }

            column = column * 26 + (upper - 'A' + 1);
        }

        return column - 1;
    }

    // The A1-style label for this address, e.g. "B3".
    public string ToA1() => $"{ColumnToLetters(this.Column)}{this.Row + 1}";

    public override string ToString() => this.ToA1();
}
