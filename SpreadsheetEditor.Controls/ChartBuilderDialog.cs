namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.Localization;

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
        ChartKind[] chartKinds = Enum.GetValues<ChartKind>();
        ComboBox typeBox = new()
        {
            ItemsSource = chartKinds.Select(kind => EditorText.Get(ChartKindKey(kind))).ToArray(),
            SelectedIndex = Array.IndexOf(chartKinds, initial?.Kind ?? ChartKind.Bar),
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        TextBox rangeBox = new() { Text = initial?.Range ?? defaultRange, PlaceholderText = EditorText.Get(EditorString.RangeExample) };
        TextBox titleBox = new() { Text = initial?.Title, PlaceholderText = EditorText.Get(EditorString.ChartTitleOptional) };
        CheckBox headersChk = new() { Content = EditorText.Get(EditorString.FirstRowHeaders), IsChecked = initial?.FirstRowHeaders ?? true };
        CheckBox catsChk = new() { Content = EditorText.Get(EditorString.FirstColumnCategories), IsChecked = initial?.FirstColCategories ?? true };
        CheckBox legendChk = new() { Content = EditorText.Get(EditorString.ShowLegend), IsChecked = initial?.ShowLegend ?? true };
        CheckBox labelsChk = new() { Content = EditorText.Get(EditorString.ShowDataLabels), IsChecked = initial?.ShowDataLabels ?? false };

        AxisNumberFormat[] axisFormats = Enum.GetValues<AxisNumberFormat>();
        ComboBox formatBox = new()
        {
            ItemsSource = axisFormats.Select(format => EditorText.Get(AxisFormatKey(format))).ToArray(),
            SelectedIndex = Array.IndexOf(axisFormats, initial?.ValueAxisFormat ?? AxisNumberFormat.Number),
            HorizontalAlignment = HorizontalAlignment.Stretch,
        };
        TextBox minBox = new() { Text = initial?.AxisMinimum?.ToString(CultureInfo.InvariantCulture), PlaceholderText = EditorText.Get(EditorString.MinimumAuto) };
        TextBox maxBox = new() { Text = initial?.AxisMaximum?.ToString(CultureInfo.InvariantCulture), PlaceholderText = EditorText.Get(EditorString.MaximumAuto) };
        TextBox stepBox = new() { Text = initial?.AxisMajorStep?.ToString(CultureInfo.InvariantCulture), PlaceholderText = EditorText.Get(EditorString.StepAuto) };
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

        Button create = new() { Content = EditorText.Get(editing ? EditorString.Save : EditorString.Create), IsDefault = true, HorizontalAlignment = HorizontalAlignment.Right };
        Button cancel = new() { Content = EditorText.Get(EditorString.Cancel), IsCancel = true };
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

        panel.Children.Add(Label(EditorString.ChartType));
        panel.Children.Add(typeBox);
        if (allowDataOptions)
        {
            panel.Children.Add(Label(EditorString.DataRange));
            panel.Children.Add(rangeBox);
        }
        else if (seriesSources is { Count: > 0 })
        {
            panel.Children.Add(Label(EditorString.DataSource));
            foreach ((string title, string range) in seriesSources)
            {
                TextBox box = new() { Text = range, PlaceholderText = EditorText.Get(EditorString.SeriesRangeExample) };
                seriesRangeBoxes.Add(box);
                panel.Children.Add(new TextBlock { Text = string.IsNullOrWhiteSpace(title) ? EditorText.Get(EditorString.SeriesValues) : title, Foreground = Brushes.Gray, Margin = new Thickness(0, 2, 0, 0) });
                panel.Children.Add(box);
            }

            categoryRangeBox = new TextBox { Text = categorySource ?? string.Empty, PlaceholderText = EditorText.Get(EditorString.CategoriesRangeExample) };
            panel.Children.Add(new TextBlock { Text = EditorText.Get(EditorString.Categories), Foreground = Brushes.Gray, Margin = new Thickness(0, 2, 0, 0) });
            panel.Children.Add(categoryRangeBox);
        }

        panel.Children.Add(Label(EditorString.Title));
        panel.Children.Add(titleBox);
        if (allowDataOptions)
        {
            panel.Children.Add(headersChk);
            panel.Children.Add(catsChk);
        }

        panel.Children.Add(legendChk);
        panel.Children.Add(labelsChk);
        panel.Children.Add(Label(EditorString.ValueAxisFormat));
        panel.Children.Add(formatBox);
        panel.Children.Add(Label(EditorString.ValueAxisScale));
        panel.Children.Add(scaleRow);
        panel.Children.Add(error);
        panel.Children.Add(buttons);

        Window dialog = new()
        {
            Title = EditorText.Get(editing ? EditorString.EditChart : EditorString.InsertChart),
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
                error.Text = EditorText.Get(EditorString.RangeInvalid);
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
                        error.Text = EditorText.Get(EditorString.SeriesRangesInvalid);
                        error.IsVisible = true;
                        return;
                    }
                }
            }

            dialog.Close(new ChartBuilderResult(
                chartKinds[Math.Max(0, typeBox.SelectedIndex)],
                rangeBox.Text ?? string.Empty,
                titleBox.Text,
                headersChk.IsChecked == true,
                catsChk.IsChecked == true,
                legendChk.IsChecked == true,
                labelsChk.IsChecked == true,
                axisFormats[Math.Max(0, formatBox.SelectedIndex)],
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

    private static TextBlock Label(EditorString key) =>
        new() { Text = EditorText.Get(key), FontWeight = FontWeight.SemiBold, Margin = new Thickness(0, 4, 0, 0) };

    private static EditorString ChartKindKey(ChartKind kind) => kind switch
    {
        ChartKind.Line => EditorString.ChartLine,
        ChartKind.Spline => EditorString.ChartSpline,
        ChartKind.StepLine => EditorString.ChartStepLine,
        ChartKind.Area => EditorString.ChartArea,
        ChartKind.StackedArea => EditorString.ChartStackedArea,
        ChartKind.Bar => EditorString.ChartBar,
        ChartKind.StackedBar => EditorString.ChartStackedBar,
        ChartKind.Scatter => EditorString.ChartScatter,
        ChartKind.Pie => EditorString.ChartPie,
        ChartKind.Donut => EditorString.ChartDonut,
        ChartKind.Radar => EditorString.ChartRadar,
        _ => EditorString.ChartBar,
    };

    private static EditorString AxisFormatKey(AxisNumberFormat format) => format switch
    {
        AxisNumberFormat.Currency => EditorString.AxisCurrency,
        AxisNumberFormat.Percent => EditorString.AxisPercent,
        _ => EditorString.AxisNumber,
    };
}
