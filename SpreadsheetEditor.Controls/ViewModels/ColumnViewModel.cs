namespace SpreadsheetEditor.Controls.ViewModels;

using CommunityToolkit.Mvvm.ComponentModel;
using global::SpreadsheetEditor.Controls.Models;

// Metadata for one spreadsheet column: its zero-based index, A/B/C header, and width.
public partial class ColumnViewModel : ObservableObject
{
    [ObservableProperty]
    private int index;

    [ObservableProperty]
    private double width = 90;

    public string Header => CellAddress.ColumnToLetters(this.Index);

    public ColumnViewModel(int index, double width)
    {
        this.index = index;
        this.width = width;
    }
}
