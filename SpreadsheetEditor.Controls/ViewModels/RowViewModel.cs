namespace SpreadsheetEditor.Controls.ViewModels;

using Avalonia.Collections;
using CommunityToolkit.Mvvm.ComponentModel;

// One spreadsheet row. Cells is positional so a generated DataGrid column at
// index i binds to Cells[i].Text.
public partial class RowViewModel : ObservableObject
{
    // One-based row number shown in the row header.
    [ObservableProperty]
    private int number;

    // Explicit row height in px (from a user drag); null = auto (max cell font + padding).
    [ObservableProperty]
    private double? height;

    // View-only filter state. Hidden rows retain their cells and original row numbers.
    [ObservableProperty]
    private bool isHidden;

    public AvaloniaList<CellViewModel> Cells { get; } = new();

    public RowViewModel(int number, int columnCount)
    {
        this.number = number;

        CellViewModel[] cells = new CellViewModel[columnCount];
        for (int c = 0; c < columnCount; c++)
        {
            cells[c] = new CellViewModel();
        }

        this.Cells.AddRange(cells);
    }
}
