namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Data;
using Avalonia.Automation;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Threading;
using global::SpreadsheetEditor.Controls.Localization;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.ViewModels;
using S = global::SpreadsheetEditor.Controls.Localization.EditorString;
using static global::SpreadsheetEditor.Controls.Localization.EditorText;

public partial class SpreadsheetGrid
{
    private readonly SolidColorBrush borderColorBrush = new(Colors.Black);

    public Color BorderColor
    {
        get => this.borderColorBrush.Color;
        set => this.borderColorBrush.Color = value;
    }
    public string BorderLineStyle { get; private set; } = "thin";

    private void SyncSelectionBorderSettings()
    {
        if (this.sheet is not { } sheet) return;
        bool first = true, sameColor = true, sameStyle = true;
        Color color = default;
        string? style = null;
        foreach (CellEdge edge in sheet.ExistingBorderEdges(this.SelectionRange))
        {
            if (first) { color = edge.Color; style = edge.Style; first = false; }
            else { sameColor &= color == edge.Color; sameStyle &= style == edge.Style; }
        }
        if (first) return;
        if (sameColor) this.BorderColor = color;
        if (sameStyle && style is not null) this.BorderLineStyle = style;
    }

    private bool UpdateSelectedBorderFormatting(SpreadsheetViewModel sheet, CellRange range, Color? color, string? style, bool checkpoint)
    {
        if (this.ReadOnly || !ReferenceEquals(this.sheet, sheet) || this.SelectionRange != range) return false;
        bool changed = sheet.SetExistingBorderFormatting(range, color, style, checkpoint);
        if (changed)
        {
            this.Rebuild();
            sheet.NotifyMetadataChanged();
        }
        return changed;
    }

    private static readonly (S Label, string Style)[] BorderStyles =
    {
        (S.BorderThin, "thin"), (S.BorderMedium, "medium"), (S.BorderThick, "thick"),
        (S.BorderDouble, "double"), (S.BorderDashed, "dashed"), (S.BorderDotted, "dotted"),
    };

    private static readonly S[] BorderLabels =
    {
        S.NoBorders, S.OutsideBorders, S.AllBorders, S.InsideBorders,
        S.TopBorder, S.BottomBorder, S.LeftBorder, S.RightBorder,
    };

    private CellAppearance? AdjacentAppearance(int row, int column)
    {
        if (this.sheet is not { } sheet || !sheet.IsValid(new CellAddress(row, column))) return null;
        CellAppearance? appearance = sheet.GetCell(new CellAddress(row, column))?.Appearance;
        if (this.covered?[row, column] != true) return appearance;
        foreach (MergeRange merge in sheet.Merges)
        {
            if (row < merge.Row || row >= merge.Row + merge.RowSpan || column < merge.Column || column >= merge.Column + merge.ColumnSpan) continue;
            CellAppearance? anchor = sheet.GetCell(new CellAddress(merge.Row, merge.Column))?.Appearance;
            return (anchor ?? new CellAppearance()) with
            {
                Right = sheet.GetCell(new CellAddress(merge.Row, merge.Column + merge.ColumnSpan - 1))?.Appearance?.Right ?? anchor?.Right,
                Bottom = sheet.GetCell(new CellAddress(merge.Row + merge.RowSpan - 1, merge.Column))?.Appearance?.Bottom ?? anchor?.Bottom,
            };
        }
        return appearance;
    }

    public async Task MergeSelectionAsync(CellMergeMode mode)
    {
        if (this.ReadOnly || this.sheet is not { } sheet || !sheet.CanMergeCells(this.SelectionRange)) return;
        this.SuspendEditor();
        CellRange range = sheet.ExpandFormattingRange(this.SelectionRange);
        if (sheet.MergeWouldDiscardData(range, mode))
        {
            if (TopLevel.GetTopLevel(this) is not Window owner) return;
            bool confirmed = await this.ConfirmMergeAsync(owner);
            if (!confirmed || this.ReadOnly || !ReferenceEquals(this.sheet, sheet) || range != sheet.ExpandFormattingRange(this.SelectionRange)) return;
        }
        this.ApplyCellUpdates(() => sheet.MergeCells(range, mode, discardContents: true));
        this.Rebuild();
        this.FocusForSelection();
    }

    public void UnmergeSelection()
    {
        if (this.ReadOnly || this.sheet is not { } sheet) return;
        this.SuspendEditor();
        if (sheet.UnmergeCells(this.SelectionRange)) this.Rebuild();
        this.FocusForSelection();
    }

    public void ApplySelectionBorders(CellBorderPreset preset)
    {
        if (this.ReadOnly || this.sheet is not { } sheet) return;
        this.SuspendEditor();
        this.ApplyCellUpdates(() => sheet.ApplyBorders(this.SelectionRange, preset, new CellEdge(this.BorderLineStyle, this.BorderColor)));
        this.Rebuild();
        this.FocusForSelection();
    }

    private async Task<bool> ConfirmMergeAsync(Window owner)
    {
        Button merge = new() { Content = Get(S.MergeCells), IsDefault = true, MinWidth = 90 };
        Button cancel = new() { Content = Get(S.Cancel), IsCancel = true, MinWidth = 90 };
        StackPanel buttons = new() { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8 };
        buttons.Children.Add(cancel);
        buttons.Children.Add(merge);
        StackPanel content = new() { Margin = new Thickness(16), Spacing = 16 };
        content.Children.Add(new TextBlock { Text = Get(S.MergeContentsWarning), TextWrapping = TextWrapping.Wrap });
        content.Children.Add(buttons);
        Window dialog = new() { Title = Get(S.MergeCells), Width = 440, SizeToContent = SizeToContent.Height,
            CanResize = false, WindowStartupLocation = WindowStartupLocation.CenterOwner, Content = content };
        merge.Click += (_, _) => dialog.Close(true);
        cancel.Click += (_, _) => dialog.Close(false);
        return await dialog.ShowDialog<bool>(owner);
    }

    internal MenuItem CreateMergeMenu() => new() { Header = Get(S.MergeCells), ItemsSource = this.CreateMergeItems() };

    internal Control[] CreateMergeItems()
    {
        List<Control> items = new();
        foreach ((S label, CellMergeMode mode) in new[]
            { (S.MergeCenter, CellMergeMode.Center), (S.MergeAcross, CellMergeMode.Across), (S.MergeCells, CellMergeMode.Merge) })
        {
            MenuItem item = new() { Header = Get(label) };
            item.Click += async (_, _) => await this.MergeSelectionAsync(mode);
            items.Add(item);
        }
        MenuItem unmerge = new() { Header = Get(S.UnmergeCells) };
        unmerge.Click += (_, _) => this.UnmergeSelection();
        items.Add(unmerge);
        return items.ToArray();
    }

    internal MenuItem CreateBordersMenu()
    {
        MenuItem menu = new() { Header = Get(S.Borders), ItemsSource = this.CreateBorderItems() };
        menu.PropertyChanged += (_, args) =>
        {
            if (args.Property == MenuItem.IsSubMenuOpenProperty && menu.IsSubMenuOpen) this.SyncSelectionBorderSettings();
        };
        return menu;
    }

    internal MenuFlyout CreateBorderFlyout()
    {
        MenuFlyout menu = new() { ItemsSource = this.CreateBorderItems() };
        menu.Opening += (_, _) => this.SyncSelectionBorderSettings();
        return menu;
    }

    internal Control[] CreateBorderItems()
    {
        List<Control> items = new();
        CellBorderPreset[] order =
        {
            CellBorderPreset.Bottom, CellBorderPreset.Top, CellBorderPreset.Left, CellBorderPreset.Right,
            CellBorderPreset.None, CellBorderPreset.All, CellBorderPreset.Outside, CellBorderPreset.Inside,
        };
        foreach (CellBorderPreset preset in order)
        {
            if (preset == CellBorderPreset.None) items.Add(new Separator());
            MenuItem item = new() { Header = Get(BorderLabels[(int)preset]) };
            item.Icon = CreateBorderIcon(item, preset);
            item.Click += (_, _) => this.ApplySelectionBorders(preset);
            items.Add(item);
        }
        items.Add(new Separator());
        MenuItem color = new()
        {
            Header = Get(S.BorderColor),
            Icon = new Border { Width = 16, Height = 16, Background = this.borderColorBrush, BorderBrush = Brushes.Gray, BorderThickness = new Thickness(1) },
        };
        MenuItem colorHost = new() { StaysOpenOnClick = true, Focusable = false };
        color.Items.Add(colorHost);
        ColorView? picker = null;
        SpreadsheetViewModel? colorSheet = null;
        CellRange colorRange = default;
        bool colorCheckpoint = false, syncing = false, colorQueued = false;
        Color? pendingColor = null;
        void ApplyPendingColor()
        {
            colorQueued = false;
            Color? value = pendingColor;
            pendingColor = null;
            if (value is not { } chosen || colorSheet is null) return;
            colorCheckpoint |= this.UpdateSelectedBorderFormatting(colorSheet, colorRange, chosen, null, checkpoint: !colorCheckpoint);
        }
        color.PropertyChanged += (_, args) =>
        {
            if (args.Property != MenuItem.IsSubMenuOpenProperty) return;
            if (!color.IsSubMenuOpen)
            {
                ApplyPendingColor();
                colorSheet = null;
                return;
            }
            colorSheet = this.sheet;
            colorRange = this.SelectionRange;
            colorCheckpoint = false;
            pendingColor = null;
            syncing = true;
            if (picker is null)
            {
                picker = EditorColorPicker.CreateView(this.BorderColor);
                AutomationProperties.SetName(picker, Get(S.BorderColor));
                picker.ColorChanged += (_, e) =>
                {
                    if (syncing || !color.IsSubMenuOpen || colorSheet is null || this.ReadOnly) return;
                    this.BorderColor = e.NewColor;
                    pendingColor = e.NewColor;
                    // Coalesce spectrum/slider input, and keep the whole picker session in one undo point.
                    if (colorQueued) return;
                    colorQueued = true;
                    Dispatcher.UIThread.Post(ApplyPendingColor, DispatcherPriority.Background);
                };
                colorHost.Header = picker;
            }
            else picker.Color = this.BorderColor;
            syncing = false;
        };
        items.Add(color);
        MenuItem styles = new() { Header = Get(S.BorderStyle) };
        string group = Guid.NewGuid().ToString();
        foreach ((S label, string style) in BorderStyles)
        {
            MenuItem item = new()
            {
                ToggleType = MenuItemToggleType.Radio,
                GroupName = group,
                IsChecked = style == this.BorderLineStyle,
                Tag = style,
            };
            item.Header = CreateLinePreview(item, style);
            ToolTip.SetTip(item, Get(label));
            AutomationProperties.SetName(item, Get(label));
            item.Click += (_, _) =>
            {
                this.BorderLineStyle = style;
                if (this.sheet is { } sheet) this.UpdateSelectedBorderFormatting(sheet, this.SelectionRange, null, style, checkpoint: true);
            };
            styles.Items.Add(item);
        }
        styles.PropertyChanged += (_, args) =>
        {
            if (args.Property == MenuItem.IsSubMenuOpenProperty && styles.IsSubMenuOpen)
                foreach (MenuItem item in styles.Items.OfType<MenuItem>()) item.IsChecked = Equals(item.Tag, this.BorderLineStyle);
        };
        items.Add(styles);
        return items.ToArray();
    }

    private static Control CreateBorderIcon(MenuItem owner, CellBorderPreset preset)
    {
        Canvas icon = new() { Width = 18, Height = 18, IsHitTestVisible = false };
        void Add(double x1, double y1, double x2, double y2, bool selected)
        {
            Line line = new()
            {
                StartPoint = new Point(x1, y1), EndPoint = new Point(x2, y2),
                StrokeThickness = selected ? 1.6 : 1,
                Opacity = selected ? 1 : 0.35,
                IsHitTestVisible = false,
            };
            if (!selected) line.StrokeDashArray = new Avalonia.Collections.AvaloniaList<double> { 1, 2 };
            line.Bind(Shape.StrokeProperty, new Binding(nameof(MenuItem.Foreground)) { Source = owner });
            icon.Children.Add(line);
        }
        bool outside = preset is CellBorderPreset.All or CellBorderPreset.Outside;
        bool inside = preset is CellBorderPreset.All or CellBorderPreset.Inside;
        Add(2, 2, 16, 2, outside || preset == CellBorderPreset.Top);
        Add(2, 16, 16, 16, outside || preset == CellBorderPreset.Bottom);
        Add(2, 2, 2, 16, outside || preset == CellBorderPreset.Left);
        Add(16, 2, 16, 16, outside || preset == CellBorderPreset.Right);
        Add(9, 2, 9, 16, inside);
        Add(2, 9, 16, 9, inside);
        return icon;
    }

    private static Control CreateLinePreview(Control owner, string style)
    {
        Canvas sample = new() { Width = 80, Height = 18, IsHitTestVisible = false };
        double thickness = style == "thick" ? 3 : style == "medium" ? 2 : 1;
        foreach (double y in style == "double" ? new[] { 6d, 10d } : new[] { 8d })
        {
            Line line = new()
            {
                StartPoint = new Point(2, y), EndPoint = new Point(78, y), StrokeThickness = thickness,
                IsHitTestVisible = false,
            };
            if (style is "dashed" or "dotted")
                line.StrokeDashArray = style == "dashed"
                    ? new Avalonia.Collections.AvaloniaList<double> { 4, 2 }
                    : new Avalonia.Collections.AvaloniaList<double> { 1, 2 };
            line.Bind(Shape.StrokeProperty, new Binding("Foreground") { Source = owner });
            sample.Children.Add(line);
        }
        return sample;
    }
}
