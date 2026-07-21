namespace SpreadsheetEditor.Controls;

using System;
using System.Globalization;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Models;

// The choices a user makes in the chart builder.
public sealed record ChartBuilderResult(
    ChartKind Kind, string Range, string? Title,
    bool FirstRowHeaders, bool FirstColCategories, bool ShowLegend, bool ShowDataLabels,
    AxisNumberFormat ValueAxisFormat = AxisNumberFormat.Number,
    double? AxisMinimum = null, double? AxisMaximum = null, double? AxisMajorStep = null);

// A small modal dialog that lets the user build a chart: type, data range, title, and a few
// options. Built in code so it needs no extra XAML. Returns null when cancelled.
public static class ChartBuilderDialog
{
    public static Task<ChartBuilderResult?> ShowAsync(Window owner, string defaultRange, ChartBuilderResult? initial = null)
    {
        bool editing = initial is not null;
        ComboBox typeBox = new()
        {
            ItemsSource = Enum.GetValues<ChartKind>(),
            SelectedItem = initial?.Kind ?? ChartKind.Bar,
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        TextBox rangeBox = new() { Text = initial?.Range ?? defaultRange, PlaceholderText = "e.g. B6:G10" };
        TextBox titleBox = new() { Text = initial?.Title, PlaceholderText = "Chart title (optional)" };
        CheckBox headersChk = new() { Content = "First row is series headers", IsChecked = initial?.FirstRowHeaders ?? true };
        CheckBox catsChk = new() { Content = "First column is category labels", IsChecked = initial?.FirstColCategories ?? true };
        CheckBox legendChk = new() { Content = "Show legend", IsChecked = initial?.ShowLegend ?? true };
        CheckBox labelsChk = new() { Content = "Show data labels", IsChecked = initial?.ShowDataLabels ?? false };

        ComboBox formatBox = new()
        {
            ItemsSource = Enum.GetValues<AxisNumberFormat>(),
            SelectedItem = initial?.ValueAxisFormat ?? AxisNumberFormat.Number,
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        TextBox minBox = new() { Text = initial?.AxisMinimum?.ToString(CultureInfo.InvariantCulture), PlaceholderText = "Min (auto)" };
        TextBox maxBox = new() { Text = initial?.AxisMaximum?.ToString(CultureInfo.InvariantCulture), PlaceholderText = "Max (auto)" };
        TextBox stepBox = new() { Text = initial?.AxisMajorStep?.ToString(CultureInfo.InvariantCulture), PlaceholderText = "Step (auto)" };
        StackPanel scaleRow = new()
        {
            Orientation = Orientation.Horizontal,
            Spacing = 6,
            Children =
            {
                new StackPanel { Width = 96, Children = { minBox } },
                new StackPanel { Width = 96, Children = { maxBox } },
                new StackPanel { Width = 96, Children = { stepBox } },
            },
        };

        TextBlock error = new() { Foreground = Brushes.IndianRed, IsVisible = false, TextWrapping = TextWrapping.Wrap };

        Button create = new() { Content = editing ? "Save" : "Create", IsDefault = true, HorizontalAlignment = HorizontalAlignment.Right };
        Button cancel = new() { Content = "Cancel", IsCancel = true };
        StackPanel buttons = new()
        {
            Orientation = Orientation.Horizontal,
            HorizontalAlignment = HorizontalAlignment.Right,
            Spacing = 8,
            Children = { cancel, create },
        };

        StackPanel panel = new()
        {
            Spacing = 6,
            Margin = new Thickness(16),
            Children =
            {
                Label("Chart type"), typeBox,
                Label("Data range"), rangeBox,
                Label("Title"), titleBox,
                headersChk, catsChk, legendChk, labelsChk,
                Label("Value axis format"), formatBox,
                Label("Value axis scale"), scaleRow,
                error,
                buttons,
            },
        };

        Window dialog = new()
        {
            Title = editing ? "Edit chart" : "Insert chart",
            Width = 340,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Content = panel,
        };

        create.Click += (_, _) =>
        {
            if (!CellRange.TryParse(rangeBox.Text, out _))
            {
                error.Text = "Enter a valid range like B6:G10.";
                error.IsVisible = true;
                return;
            }

            dialog.Close(new ChartBuilderResult(
                (ChartKind)(typeBox.SelectedItem ?? ChartKind.Bar),
                rangeBox.Text ?? string.Empty,
                titleBox.Text,
                headersChk.IsChecked == true,
                catsChk.IsChecked == true,
                legendChk.IsChecked == true,
                labelsChk.IsChecked == true,
                (AxisNumberFormat)(formatBox.SelectedItem ?? AxisNumberFormat.Number),
                ParseOptional(minBox.Text),
                ParseOptional(maxBox.Text),
                ParseOptional(stepBox.Text)));
        };
        cancel.Click += (_, _) => dialog.Close(null);

        return dialog.ShowDialog<ChartBuilderResult?>(owner);
    }

    // Parses an optional axis bound; blank/invalid → null (auto).
    private static double? ParseOptional(string? text) =>
        double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double v) ? v : null;

    private static TextBlock Label(string text) =>
        new() { Text = text, FontWeight = FontWeight.SemiBold, Margin = new Thickness(0, 4, 0, 0) };
}
