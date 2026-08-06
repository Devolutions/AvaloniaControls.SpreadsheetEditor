namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Models;

// The choices a user makes in the chart builder. SeriesRanges/CategoryRange are populated only when
// editing an imported chart's explicit source ranges (one A1 range per series, plus the categories).
public sealed record ChartBuilderResult(
    ChartKind Kind, string Range, string? Title,
    bool FirstRowHeaders, bool FirstColCategories, bool ShowLegend, bool ShowDataLabels,
    AxisNumberFormat ValueAxisFormat = AxisNumberFormat.Number,
    double? AxisMinimum = null, double? AxisMaximum = null, double? AxisMajorStep = null,
    IReadOnlyList<string>? SeriesRanges = null, string? CategoryRange = null);

// A small modal dialog that lets the user build a chart: type, data range, title, and a few
// options. Built in code so it needs no extra XAML. Returns null when cancelled.
public static class ChartBuilderDialog
{
    // allowDataOptions: whether the block data-range + first-row/first-column options are shown (charts
    // built from a contiguous cell block). seriesSources/categorySource: for an imported chart, the
    // per-series value ranges (title + A1) and category range are shown as editable fields so the user
    // can retarget where the chart's values come from without a single block.
    public static Task<ChartBuilderResult?> ShowAsync(
        Window owner, string defaultRange, ChartBuilderResult? initial = null, bool allowDataOptions = true,
        IReadOnlyList<(string Title, string Range)>? seriesSources = null, string? categorySource = null)
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

        StackPanel panel = new() { Spacing = 6, Margin = new Thickness(16) };
        // Editable value range per series + the category range (for imported charts with explicit series).
        List<TextBox> seriesRangeBoxes = new();
        TextBox? categoryRangeBox = null;

        panel.Children.Add(Label("Chart type"));
        panel.Children.Add(typeBox);
        if (allowDataOptions)
        {
            panel.Children.Add(Label("Data range"));
            panel.Children.Add(rangeBox);
        }
        else if (seriesSources is { Count: > 0 })
        {
            panel.Children.Add(Label("Data source"));
            foreach ((string title, string range) in seriesSources)
            {
                TextBox box = new() { Text = range, PlaceholderText = "e.g. G9:G10 or C7, C10" };
                seriesRangeBoxes.Add(box);
                panel.Children.Add(new TextBlock { Text = string.IsNullOrWhiteSpace(title) ? "Values" : title, Foreground = Brushes.Gray, Margin = new Thickness(0, 2, 0, 0) });
                panel.Children.Add(box);
            }

            categoryRangeBox = new TextBox { Text = categorySource ?? string.Empty, PlaceholderText = "Range (E9:E10) or names (Income, Expenses)" };
            panel.Children.Add(new TextBlock { Text = "Categories", Foreground = Brushes.Gray, Margin = new Thickness(0, 2, 0, 0) });
            panel.Children.Add(categoryRangeBox);
        }

        panel.Children.Add(Label("Title"));
        panel.Children.Add(titleBox);
        if (allowDataOptions)
        {
            panel.Children.Add(headersChk);
            panel.Children.Add(catsChk);
        }

        panel.Children.Add(legendChk);
        panel.Children.Add(labelsChk);
        panel.Children.Add(Label("Value axis format"));
        panel.Children.Add(formatBox);
        panel.Children.Add(Label("Value axis scale"));
        panel.Children.Add(scaleRow);
        panel.Children.Add(error);
        panel.Children.Add(buttons);

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
            if (allowDataOptions && !CellRange.TryParse(rangeBox.Text, out _))
            {
                error.Text = "Enter a valid range like B6:G10.";
                error.IsVisible = true;
                return;
            }

            // Validate each series' value ranges (blank keeps the existing range). A series can list
            // several comma/semicolon-separated cells/ranges (e.g. "C7, C10"); each must be a valid range.
            foreach (TextBox box in seriesRangeBoxes)
            {
                if (string.IsNullOrWhiteSpace(box.Text))
                {
                    continue;
                }

                foreach (string token in box.Text.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    if (!CellRange.TryParse(token.Trim(), out _))
                    {
                        error.Text = "Enter valid value cells/ranges like G9:G10 or C7, C10.";
                        error.IsVisible = true;
                        return;
                    }
                }
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
                ParseOptional(stepBox.Text),
                seriesRangeBoxes.Count > 0 ? seriesRangeBoxes.ConvertAll(b => b.Text ?? string.Empty) : null,
                categoryRangeBox?.Text));
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
