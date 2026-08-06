namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Primitives;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Reactive;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.Printing;
using global::SpreadsheetEditor.Controls.ViewModels;

// A custom spreadsheet grid: a scrollable cell area with frozen column/row headers, per-column
// widths and per-row heights, fills, inline-run formatting, and merged-cell spanning. Replaces the
// Avalonia DataGrid so cells can span (merge) and (later) host a rich editor. Cell content/editing
// state lives in SpreadsheetViewModel; this control only renders and raises CellActivated on click.
public class SpreadsheetGrid : UserControl
{
    private const double RowHeaderWidth = 46;
    private const double ColumnHeaderHeight = 22;
    private const double DefaultRowHeight = 22;

    // Theme-agnostic accents (blue selection / teal active header read fine on both light and dark).
    private static readonly IBrush ActiveHeaderBackground = new SolidColorBrush(Color.FromRgb(0x57, 0xC7, 0xCE));
    private static readonly IBrush SelectionBorder = new SolidColorBrush(Color.FromRgb(0x1A, 0x73, 0xE8));
    private static readonly IBrush AnchorBorder = new SolidColorBrush(Color.FromRgb(0x7F, 0xB0, 0xEE)); // lighter than SelectionBorder
    private static readonly IBrush SelectionFillBrush = new SolidColorBrush(Color.FromArgb(0x33, 0x1A, 0x73, 0xE8));

    // Theme-dependent brushes (headers, gridlines, sheet surface). Initialized to the light palette and
    // re-resolved from the active ThemeVariant in ApplyTheme() (on attach + on theme change).
    private IBrush headerBackground = new SolidColorBrush(Color.FromRgb(0xF3, 0xF3, 0xF3));
    private IBrush headerText = new SolidColorBrush(Color.FromRgb(0x60, 0x60, 0x60));
    private IBrush gridLine = new SolidColorBrush(Color.FromArgb(0x30, 0, 0, 0));
    private IBrush surface = Brushes.White;

    private readonly ScrollViewer cellScroller = new()
    { HorizontalScrollBarVisibility = ScrollBarVisibility.Auto, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };

    private readonly ScrollViewer columnHeaderScroller = new()
    { HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden, VerticalScrollBarVisibility = ScrollBarVisibility.Disabled };

    private readonly ScrollViewer rowHeaderScroller = new()
    { HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, VerticalScrollBarVisibility = ScrollBarVisibility.Hidden };

    private readonly Grid columnHeaders = new();
    private readonly Grid rowHeaders = new();
    private readonly Grid cells = new();
    private readonly Border selection = new()
    {
        BorderBrush = SelectionBorder,
        BorderThickness = new Thickness(2),
        IsHitTestVisible = false,
        IsVisible = false,
    };

    // Translucent fill over a multi-cell selected range, in up to 4 parts so the anchor cell is left
    // unfilled (Excel-style). The outer rectangle is selection; anchorBorder outlines the anchor.
    private readonly Border[] fillParts =
    {
        NewFill(), NewFill(), NewFill(), NewFill(),
    };

    private readonly Border anchorBorder = new()
    {
        BorderBrush = AnchorBorder,
        BorderThickness = new Thickness(1),
        IsHitTestVisible = false,
        IsVisible = false,
    };

    private static Border NewFill() => new()
    {
        Background = SelectionFillBrush,
        IsHitTestVisible = false,
        IsVisible = false,
    };

    // Selection range: anchor (the active/editable cell) to focus (the drag/extend end).
    private CellAddress rangeAnchor;
    private CellAddress rangeFocus;
    private bool rangeDragging;
    private bool chartDragging;
    private SheetChart? selectedChart;   // the chart targeted by the Delete key (null = none)

    // Internal clipboard (rich CellData block) + the TSV we last wrote to the OS clipboard, so an
    // internal paste keeps formatting/formulas while an external (Excel) copy still pastes as values.
    private CellData[][]? clipboard;
    private string? clipboardText;

    // A single in-cell editor moved over the active cell (rather than one editor per cell). It hosts
    // the Pro RichTextEditor (per-character formatting) with a plain-TextBox fallback. Wrapped in an
    // opaque Border so it covers the display text beneath it.
    private readonly CellEditor editor = new();
    private readonly Border editorHost;
    private readonly Border cornerBox;

    // Embedded pictures, drawn on a canvas above the cell fills but below the selection/editor.
    private readonly Canvas imageOverlay = new() { IsHitTestVisible = false };

    // Charts, on their own canvas above images (hit-testable so chart tooltips/interaction work).
    private readonly Canvas chartOverlay = new();

    // Live chart cards tracked as they are built (for print/PDF snapshotting): the positioned card,
    // its chart control, and the on-card chrome (⚙ button + resize grip) hidden when snapshotting.
    private readonly List<(Border Card, Control Chart, Control Settings, Control Grip)> chartCards = new();

    // True = selection mode (editor hidden, just the selection rectangle); false = edit mode (editor
    // shown + focused). Default selection mode: a single click selects, double-click/typing/F2 edits.
    private bool editorSuspended = true;

    private bool adjustingRows;            // reentrancy guard for dynamic row growth/trim
    private bool adjustingColumns;         // reentrancy guard for dynamic column growth/trim
    private const int RowBuffer = 25;       // rows kept past the last visible/used row
    private const int ColumnBuffer = 8;     // columns kept past the last visible/used column

    private SpreadsheetViewModel? sheet;
    private double[] columnWidths = Array.Empty<double>();
    private double[] rowHeights = Array.Empty<double>();
    private bool[,]? covered;
    private (int RowSpan, int ColumnSpan)[,]? spans;
    private readonly List<(CellViewModel Cell, PropertyChangedEventHandler Handler)> cellSubscriptions = new();

    public event Action<CellAddress>? CellActivated;

    // When true, the grid blocks all content mutation (editing, typing, paste/cut/clear, formatting).
    public bool ReadOnly { get; set; }

    public SpreadsheetGrid()
    {
        // Inset by 2px so the opaque editor doesn't paint over the cell gridlines / selection border
        // (it sits just inside the selection rectangle, like Excel's active cell).
        this.editorHost = new Border { IsVisible = false, Child = this.editor, Margin = new Thickness(2) };
        this.editor.SelectionFormatChanged += (_, e) => this.SelectionFormatChanged?.Invoke(this, e);

        Grid root = new()
        {
            RowDefinitions = new RowDefinitions("Auto,*"),
            ColumnDefinitions = new ColumnDefinitions("Auto,*"),
        };

        this.cornerBox = new Border
        {
            Background = this.headerBackground,
            BorderBrush = this.gridLine,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Width = RowHeaderWidth,
            Height = ColumnHeaderHeight,
        };
        Grid.SetRow(this.cornerBox, 0); Grid.SetColumn(this.cornerBox, 0);

        this.columnHeaderScroller.Content = this.columnHeaders;
        Grid.SetRow(this.columnHeaderScroller, 0); Grid.SetColumn(this.columnHeaderScroller, 1);

        this.rowHeaderScroller.Content = this.rowHeaders;
        Grid.SetRow(this.rowHeaderScroller, 1); Grid.SetColumn(this.rowHeaderScroller, 0);

        this.cells.Children.Add(this.selection);
        this.cells.Children.Add(this.editorHost);
        this.cellScroller.Content = this.cells;
        Grid.SetRow(this.cellScroller, 1); Grid.SetColumn(this.cellScroller, 1);

        // Keep the frozen headers in sync with the cell area's scroll position.
        this.cellScroller.GetObservable(ScrollViewer.OffsetProperty).Subscribe(new AnonymousObserver<Vector>(offset =>
        {
            this.columnHeaderScroller.Offset = this.columnHeaderScroller.Offset.WithX(offset.X);
            this.rowHeaderScroller.Offset = this.rowHeaderScroller.Offset.WithY(offset.Y);
        }));

        // Dynamic "infinite" rows: grow near the bottom, trim trailing empty rows when scrolled up.
        this.cellScroller.ScrollChanged += (_, _) =>
        {
            this.AdjustRowsForScroll();
            this.AdjustColumnsForScroll();
        };

        // Right-click menu (acts on the selection).
        this.cells.ContextFlyout = this.BuildCellMenu();

        root.Children.Add(this.cornerBox);
        root.Children.Add(this.columnHeaderScroller);
        root.Children.Add(this.rowHeaderScroller);
        root.Children.Add(this.cellScroller);
        this.Content = root;

        // Intercept navigation keys before the in-cell editor (tunnel) so arrows/Enter/Tab move the
        // selection. Typing on a selected cell starts editing. The grid is focusable so it can receive
        // these keys while in selection mode (no editor focused).
        this.Focusable = true;
        this.AddHandler(KeyDownEvent, this.OnNavigationKey, RoutingStrategies.Tunnel);
        this.AddHandler(TextInputEvent, this.OnTextInput, RoutingStrategies.Tunnel);

        this.cells.Background = this.surface;
        this.cellScroller.Background = this.surface;
        // Re-resolve the palette whenever the light/dark theme changes at runtime.
        this.ActualThemeVariantChanged += (_, _) => this.ApplyTheme();
    }

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        this.ApplyTheme(); // ActualThemeVariant is resolved once attached to the tree
    }

    // Resolves the header/gridline/surface brushes from the active light/dark theme and re-renders.
    private void ApplyTheme()
    {
        bool dark = this.ActualThemeVariant == ThemeVariant.Dark;
        this.headerBackground = new SolidColorBrush(dark ? Color.FromRgb(0x2D, 0x2D, 0x30) : Color.FromRgb(0xF3, 0xF3, 0xF3));
        this.headerText = new SolidColorBrush(dark ? Color.FromRgb(0xC8, 0xC8, 0xC8) : Color.FromRgb(0x60, 0x60, 0x60));
        this.gridLine = new SolidColorBrush(dark ? Color.FromArgb(0x28, 0xFF, 0xFF, 0xFF) : Color.FromArgb(0x30, 0, 0, 0));
        this.surface = new SolidColorBrush(dark ? Color.FromRgb(0x1E, 0x1E, 0x1E) : Colors.White);

        this.cornerBox.Background = this.headerBackground;
        this.cornerBox.BorderBrush = this.gridLine;
        this.cells.Background = this.surface;
        this.cellScroller.Background = this.surface;
        this.Rebuild(); // headers + cells rebuild using the new brushes
    }

    // Arrow keys / Enter / Tab move the selected cell. F2 (or double-click / typing) enters edit
    // mode. While editing, arrows move the caret; Enter/Tab commit and move.
    private void OnNavigationKey(object? sender, KeyEventArgs e)
    {
        if (this.sheet is null || this.sheet.RowCount == 0 || this.sheet.ColumnCount == 0)
        {
            return;
        }

        // A selected chart takes the Delete/Backspace key: remove it instead of clearing cells.
        if (this.selectedChart is not null && (e.Key == Key.Delete || e.Key == Key.Back))
        {
            SheetChart toRemove = this.selectedChart;
            this.selectedChart = null;
            this.sheet.RemoveChart(toRemove);
            e.Handled = true;
            return;
        }

        // Clipboard / range shortcuts act on the selection — only in selection mode, so while editing
        // the in-cell editor keeps its own text copy/paste/select-all.
        if (this.editorSuspended)
        {
            bool ctrl = e.KeyModifiers.HasFlag(KeyModifiers.Control);
            if (ctrl)
            {
                switch (e.Key)
                {
                    case Key.C: this.CopySelection(); e.Handled = true; return;
                    case Key.X: this.CutSelection(); e.Handled = true; return;
                    case Key.V: this.PasteSelection(); e.Handled = true; return;
                    case Key.A: this.SelectAll(); e.Handled = true; return;
                }
            }
            if (e.Key == Key.Delete || e.Key == Key.Back)
            {
                this.ClearSelection();
                e.Handled = true;
                return;
            }
        }

        // F2 enters edit mode on the active cell (keeping its content).
        if (e.Key == Key.F2)
        {
            if (this.editorSuspended && !this.ReadOnly)
            {
                this.editorSuspended = false;
                this.FocusCell();
            }
            e.Handled = true;
            return;
        }

        bool shift = e.KeyModifiers.HasFlag(KeyModifiers.Shift);
        int dr = 0, dc = 0;
        bool arrow = false;
        switch (e.Key)
        {
            case Key.Up: dr = -1; arrow = true; break;
            case Key.Down: dr = 1; arrow = true; break;
            case Key.Left: dc = -1; arrow = true; break;
            case Key.Right: dc = 1; arrow = true; break;
            case Key.Enter: dr = shift ? -1 : 1; break;
            case Key.Tab: dc = shift ? -1 : 1; break;
            default: return;
        }

        // While editing, let arrow keys move the caret in the editor; only Enter/Tab commit + move.
        if (arrow && !this.editorSuspended)
        {
            return;
        }

        // Shift+arrow extends the selection from the anchor instead of moving it.
        if (arrow && shift)
        {
            this.ExtendSelectionTo(new CellAddress(this.rangeFocus.Row + dr, this.rangeFocus.Column + dc));
            this.selection.BringIntoView();
            e.Handled = true;
            return;
        }

        CellAddress current = this.sheet.ActiveCell;
        int row = Math.Clamp(current.Row + dr, 0, this.sheet.RowCount - 1);
        int col = Math.Clamp(current.Column + dc, 0, this.sheet.ColumnCount - 1);
        e.Handled = true;
        if (row == current.Row && col == current.Column)
        {
            return;
        }

        this.MoveSelection(new CellAddress(row, col));
    }

    // Moves the selection to a cell (selection mode, no editing), scrolls it into view, and keeps
    // keyboard focus on the grid for further navigation.
    private void MoveSelection(CellAddress target)
    {
        this.ClearChartSelection();
        this.editorSuspended = true;
        this.CellActivated?.Invoke(target);
        Dispatcher.UIThread.Post(() =>
        {
            this.UpdateSelection();
            this.selection.BringIntoView();
            this.Focus();
        });
    }

    // Extends the selection's focus end to a cell (keeping the anchor) — drag/shift-select.
    private void ExtendSelectionTo(CellAddress focus)
    {
        if (this.sheet is null)
        {
            return;
        }
        this.rangeFocus = new CellAddress(
            Math.Clamp(focus.Row, 0, this.sheet.RowCount - 1),
            Math.Clamp(focus.Column, 0, this.sheet.ColumnCount - 1));
        this.editorSuspended = true;
        this.UpdateSelection();
        this.Focus();
    }

    // The currently selected rectangular range (anchor to focus).
    public CellRange SelectionRange => new(
        Math.Min(this.rangeAnchor.Row, this.rangeFocus.Row),
        Math.Min(this.rangeAnchor.Column, this.rangeFocus.Column),
        Math.Abs(this.rangeAnchor.Row - this.rangeFocus.Row) + 1,
        Math.Abs(this.rangeAnchor.Column - this.rangeFocus.Column) + 1);

    // The column index at a given X pixel within the cell area.
    private int ColAtX(double x)
    {
        double acc = 0;
        for (int c = 0; c < this.columnWidths.Length; c++)
        {
            acc += this.columnWidths[c];
            if (acc >= x)
            {
                return c;
            }
        }
        return Math.Max(0, this.columnWidths.Length - 1);
    }

    // Selects a chart (highlights it; the Delete key then removes it) and focuses the grid.
    private void SelectChart(SheetChart chart)
    {
        if (ReferenceEquals(this.selectedChart, chart))
        {
            this.Focus();
            return;
        }
        this.selectedChart = chart;
        this.RefreshCharts();   // re-render so the selected card shows its highlight border
        this.Focus();
    }

    // Clears the chart selection (e.g. when a cell is selected instead).
    private void ClearChartSelection()
    {
        if (this.selectedChart is null)
        {
            return;
        }
        this.selectedChart = null;
        this.RefreshCharts();
    }

    // True when node is ancestor or one of its descendants.
    private static bool IsWithin(Visual? node, Visual ancestor)
    {
        for (Visual? v = node; v is not null; v = v.GetVisualParent())
        {
            if (ReferenceEquals(v, ancestor))
            {
                return true;
            }
        }
        return false;
    }

    // The column whose left boundary is nearest the given X pixel (used to snap a dragged chart).
    private int NearestColumn(double x)
    {
        int best = 0;
        double bestDist = double.MaxValue, acc = 0;
        for (int c = 0; c <= this.columnWidths.Length; c++)
        {
            double d = Math.Abs(acc - x);
            if (d < bestDist)
            {
                bestDist = d;
                best = c;
            }
            if (c < this.columnWidths.Length)
            {
                acc += this.columnWidths[c];
            }
        }
        return Math.Min(best, Math.Max(0, this.columnWidths.Length - 1));
    }

    // The row whose top boundary is nearest the given Y pixel (used to snap a dragged chart).
    private int NearestRow(double y)
    {
        int best = 0;
        double bestDist = double.MaxValue, acc = 0;
        for (int r = 0; r <= this.rowHeights.Length; r++)
        {
            double d = Math.Abs(acc - y);
            if (d < bestDist)
            {
                bestDist = d;
                best = r;
            }
            if (r < this.rowHeights.Length)
            {
                acc += this.rowHeights[r];
            }
        }
        return Math.Min(best, Math.Max(0, this.rowHeights.Length - 1));
    }

    // Single click on a cell: selects it (selection mode), no editing.
    private void SelectOnly(CellAddress address)
    {
        this.ClearChartSelection();
        this.editorSuspended = true;
        this.CellActivated?.Invoke(address);
        Dispatcher.UIThread.Post(() =>
        {
            this.UpdateSelection();
            this.Focus();
        });
    }

    // Enters edit mode on a cell (shows + focuses the editor).
    private void BeginEdit(CellAddress address)
    {
        if (this.ReadOnly)
        {
            this.SelectOnly(address); // read-only: select but don't edit
            return;
        }

        this.sheet?.Checkpoint(); // record the pre-edit state for undo
        this.CellActivated?.Invoke(address);
        this.FocusCell(selectAll: true); // double-click to edit highlights the whole text
    }

    // Typing on a selected cell starts editing, replacing its content with the typed text.
    private void OnTextInput(object? sender, TextInputEventArgs e)
    {
        if (this.sheet is null || !this.editorSuspended || this.ReadOnly)
        {
            return; // already editing, or read-only — let the editor handle the input
        }
        if (string.IsNullOrEmpty(e.Text))
        {
            return;
        }
        CellViewModel? cell = this.sheet.GetCell(this.sheet.ActiveCell);
        if (cell is null)
        {
            return;
        }

        this.sheet.Checkpoint();  // record the pre-edit state for undo
        this.editorSuspended = false;
        cell.Text = e.Text;       // replace existing content with the typed character (Excel-style)
        this.UpdateSelection();   // shows + loads the editor with the new text
        Dispatcher.UIThread.Post(() => this.FocusEditor());
        e.Handled = true;
    }

    // Selects the active cell (selection mode) and focuses the grid — used on startup so the app
    // opens with a selected cell rather than an open editor.
    public void FocusForSelection()
    {
        this.editorSuspended = true;
        this.UpdateSelection();
        this.Focus();
    }

    public SpreadsheetViewModel? Sheet
    {
        get => this.sheet;
        set
        {
            if (this.sheet is not null)
            {
                this.sheet.PropertyChanged -= this.OnSheetPropertyChanged;
                this.sheet.DataChanged -= this.RefreshCharts;
                this.sheet.Charts.CollectionChanged -= this.OnChartsChanged;
            }
            this.sheet = value;
            if (this.sheet is not null)
            {
                this.sheet.PropertyChanged += this.OnSheetPropertyChanged;
                this.sheet.DataChanged += this.RefreshCharts;     // live-refresh charts when data recalculates
                this.sheet.Charts.CollectionChanged += this.OnChartsChanged;
            }
            this.Rebuild();
        }
    }

    private void OnChartsChanged(object? sender, System.Collections.Specialized.NotifyCollectionChangedEventArgs e) => this.RefreshCharts();

    private void OnSheetPropertyChanged(object? sender, PropertyChangedEventArgs e)
    {
        if (e.PropertyName == nameof(SpreadsheetViewModel.ActiveCell))
        {
            // A new active cell collapses the range to that single cell (drag/shift-extend set the
            // focus directly without changing ActiveCell, so they aren't reset here).
            this.rangeAnchor = this.rangeFocus = this.sheet!.ActiveCell;
            this.UpdateSelection();
        }
    }

    // Rebuilds the entire grid from the current sheet (call after structural changes / load).
    public void Rebuild()
    {
        foreach ((CellViewModel cell, PropertyChangedEventHandler handler) in this.cellSubscriptions)
        {
            cell.PropertyChanged -= handler;
        }
        this.cellSubscriptions.Clear();

        this.columnHeaders.Children.Clear();
        this.rowHeaders.Children.Clear();
        this.cells.Children.Clear();
        this.imageOverlay.Children.Clear();
        this.columnHeaders.ColumnDefinitions.Clear();
        this.columnHeaders.RowDefinitions.Clear();
        this.rowHeaders.RowDefinitions.Clear();
        this.rowHeaders.ColumnDefinitions.Clear();
        this.cells.ColumnDefinitions.Clear();
        this.cells.RowDefinitions.Clear();

        if (this.sheet is null || this.sheet.RowCount == 0 || this.sheet.ColumnCount == 0)
        {
            return;
        }

        int cols = this.sheet.ColumnCount;
        int rows = this.sheet.RowCount;

        this.columnWidths = new double[cols];
        for (int c = 0; c < cols; c++)
        {
            this.columnWidths[c] = this.sheet.Columns[c].Width > 0 ? this.sheet.Columns[c].Width : this.sheet.DefaultColumnWidth;
        }

        this.rowHeights = new double[rows];
        for (int r = 0; r < rows; r++)
        {
            if (this.sheet.Rows[r].Height is { } explicitHeight)
            {
                this.rowHeights[r] = explicitHeight;
                continue;
            }
            double maxFont = this.sheet.Rows[r].Cells.Count == 0 ? CellViewModel.DefaultFontSize : this.sheet.Rows[r].Cells.Max(c => c.FontSize);
            this.rowHeights[r] = Math.Max(DefaultRowHeight, maxFont + 8);
        }

        this.BuildMergeMaps(rows, cols);

        // Column/row definitions shared in layout across headers + cells (so they stay aligned).
        foreach (double w in this.columnWidths)
        {
            this.columnHeaders.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(w)));
            this.cells.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(w)));
        }
        this.columnHeaders.RowDefinitions.Add(new RowDefinition(new GridLength(ColumnHeaderHeight)));
        this.rowHeaders.ColumnDefinitions.Add(new ColumnDefinition(new GridLength(RowHeaderWidth)));
        foreach (double h in this.rowHeights)
        {
            this.rowHeaders.RowDefinitions.Add(new RowDefinition(new GridLength(h)));
            this.cells.RowDefinitions.Add(new RowDefinition(new GridLength(h)));
        }

        for (int c = 0; c < cols; c++)
        {
            this.columnHeaders.Children.Add(HeaderCell(CellAddress.ColumnToLetters(c), 0, c));
            this.columnHeaders.Children.Add(this.ColumnResizeHandle(c));
        }
        for (int r = 0; r < rows; r++)
        {
            this.rowHeaders.Children.Add(HeaderCell((r + 1).ToString(), r, 0));
            this.rowHeaders.Children.Add(this.RowResizeHandle(r));
        }

        List<Control> spillOverlays = new();
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols && c < this.sheet.Rows[r].Cells.Count; c++)
            {
                if (this.covered![r, c])
                {
                    continue;
                }
                (int rowSpan, int colSpan) = this.spans![r, c];
                CellViewModel cell = this.sheet.Rows[r].Cells[c];
                (int spillLeft, int spillRight) = colSpan == 1 ? this.ComputeSpill(r, c, cell) : (0, 0);
                this.cells.Children.Add(this.BuildCell(cell, r, c, rowSpan, colSpan, spilling: spillLeft + spillRight > 0));
                if (spillLeft + spillRight > 0)
                {
                    spillOverlays.Add(this.BuildSpillText(cell, r, c - spillLeft, spillLeft + 1 + spillRight));
                }
            }
        }

        // Spilling text sits above the cell grid (so it overflows empty neighbours' gridlines) but is
        // not hit-testable, so the underlying cells stay selectable.
        foreach (Control overlay in spillOverlays)
        {
            this.cells.Children.Add(overlay);
        }

        // Layer order (last = on top): cell fills/text < images < charts < selection < editor.
        this.BuildImages(rows, cols);
        Grid.SetRowSpan(this.imageOverlay, rows);
        Grid.SetColumnSpan(this.imageOverlay, cols);
        Grid.SetRowSpan(this.chartOverlay, rows);
        Grid.SetColumnSpan(this.chartOverlay, cols);
        this.cells.Children.Add(this.imageOverlay);
        this.cells.Children.Add(this.chartOverlay);
        foreach (Border part in this.fillParts)
        {
            this.cells.Children.Add(part);
        }
        this.cells.Children.Add(this.selection);
        this.cells.Children.Add(this.anchorBorder);
        this.cells.Children.Add(this.editorHost);

        this.RefreshCharts();
        this.UpdateSelection();
    }

    // Places each embedded picture on the overlay canvas, positioned from the running sum of our
    // column widths/row heights to the anchor cell, plus the picture's in-cell pixel offset.
    private void BuildImages(int rows, int cols)
    {
        if (this.sheet is null || this.sheet.Images.Count == 0)
        {
            return;
        }

        foreach (SheetImage img in this.sheet.Images)
        {
            Bitmap bitmap;
            try
            {
                using MemoryStream stream = new(img.Data);
                bitmap = new Bitmap(stream);
            }
            catch
            {
                continue; // unsupported/corrupt image data — skip it
            }

            double left = img.OffsetX;
            for (int c = 0; c < img.Column && c < cols; c++)
            {
                left += this.columnWidths[c];
            }
            double top = img.OffsetY;
            for (int r = 0; r < img.Row && r < rows; r++)
            {
                top += this.rowHeights[r];
            }

            Image control = new()
            {
                Source = bitmap,
                Width = img.Width,
                Height = img.Height,
                Stretch = Stretch.Fill,
            };
            Canvas.SetLeft(control, left);
            Canvas.SetTop(control, top);
            this.imageOverlay.Children.Add(control);
        }
    }

    private double OffsetToColumn(int column)
    {
        double x = 0;
        for (int c = 0; c < column && c < this.columnWidths.Length; c++)
        {
            x += this.columnWidths[c];
        }
        return x;
    }

    private double OffsetToRow(int row)
    {
        double y = 0;
        for (int r = 0; r < row && r < this.rowHeights.Length; r++)
        {
            y += this.rowHeights[r];
        }
        return y;
    }

    // The row index at a given Y pixel within the cell area.
    private int RowAtY(double y)
    {
        double acc = 0;
        for (int r = 0; r < this.rowHeights.Length; r++)
        {
            acc += this.rowHeights[r];
            if (acc >= y)
            {
                return r;
            }
        }
        return Math.Max(0, this.rowHeights.Length - 1);
    }

    // Grows the sheet when the user nears the bottom and trims surplus empty rows when they scroll
    // back up — an Excel-like effectively-unbounded grid without per-row virtualization.
    private void AdjustRowsForScroll()
    {
        if (this.sheet is null || this.adjustingRows || this.adjustingColumns || this.resizing || this.chartDragging || this.rowHeights.Length == 0)
        {
            return;
        }

        double viewport = this.cellScroller.Viewport.Height;
        double extent = this.cellScroller.Extent.Height;
        double offsetY = this.cellScroller.Offset.Y;
        if (viewport <= 0)
        {
            return;
        }

        int target = this.sheet.RowCount;
        if (offsetY + viewport >= extent - 40)
        {
            // At/near the bottom (or content shorter than the viewport) → add a chunk of rows.
            target = this.sheet.RowCount + RowBuffer;
        }
        else
        {
            int bottomVisibleRow = this.RowAtY(offsetY + viewport) + 1;
            int needed = Math.Max(this.sheet.LastNonEmptyRow() + 1, bottomVisibleRow) + RowBuffer;
            if (this.sheet.RowCount > needed + RowBuffer) // hysteresis so we don't churn while scrolling
            {
                target = needed;
            }
        }

        if (target == this.sheet.RowCount)
        {
            return;
        }

        this.adjustingRows = true;
        try
        {
            this.sheet.SetRowCount(target);
            this.Rebuild();
        }
        finally
        {
            this.adjustingRows = false;
        }
    }

    // Column analogue of AdjustRowsForScroll: grows the sheet when the user nears the right edge (e.g.
    // Shift+scroll to scroll horizontally) and trims surplus empty columns when they scroll back left.
    // Also fills the viewport width when the content is narrower than the visible area.
    private void AdjustColumnsForScroll()
    {
        if (this.sheet is null || this.adjustingColumns || this.adjustingRows || this.resizing || this.chartDragging || this.columnWidths.Length == 0)
        {
            return;
        }

        double viewport = this.cellScroller.Viewport.Width;
        double extent = this.cellScroller.Extent.Width;
        double offsetX = this.cellScroller.Offset.X;
        if (viewport <= 0)
        {
            return;
        }

        int target = this.sheet.ColumnCount;
        if (offsetX + viewport >= extent - 40)
        {
            // At/near the right (or content narrower than the viewport) → add a chunk of columns.
            target = this.sheet.ColumnCount + ColumnBuffer;
        }
        else
        {
            int rightVisibleColumn = this.ColAtX(offsetX + viewport) + 1;
            int needed = Math.Max(this.sheet.LastNonEmptyColumn() + 1, rightVisibleColumn) + ColumnBuffer;
            if (this.sheet.ColumnCount > needed + ColumnBuffer) // hysteresis so we don't churn while scrolling
            {
                target = needed;
            }
        }

        if (target == this.sheet.ColumnCount)
        {
            return;
        }

        this.adjustingColumns = true;
        try
        {
            this.sheet.SetColumnCount(target);
            this.Rebuild();
        }
        finally
        {
            this.adjustingColumns = false;
        }
    }

    // Rebuilds every chart overlay from the current cell data (called on structural changes, when a
    // chart is added/removed, and after each recalculation so charts stay live).
    private void RefreshCharts()
    {
        this.chartOverlay.Children.Clear();
        this.chartCards.Clear();
        if (this.sheet is null || this.columnWidths.Length == 0)
        {
            return;
        }

        // Drop a stale selection (chart deleted or sheet reloaded) so it can't highlight a missing card.
        if (this.selectedChart is not null && !this.sheet.Charts.Contains(this.selectedChart))
        {
            this.selectedChart = null;
        }

        foreach (SheetChart chart in this.sheet.Charts)
        {
            double left = this.OffsetToColumn(chart.AnchorColumn) + chart.OffsetX;
            double top = this.OffsetToRow(chart.AnchorRow) + chart.OffsetY;

            // Size to span from→to cells in OUR layout (so it covers the same cells Excel anchored it
            // to); fall back to the stored pixel size for one-cell-anchored charts.
            double width = chart.Width, height = chart.Height;
            if (chart.ToColumn is { } toCol && chart.ToRow is { } toRow)
            {
                width = this.OffsetToColumn(toCol) + chart.ToOffsetX - left;
                height = this.OffsetToRow(toRow) + chart.ToOffsetY - top;
            }
            width = Math.Clamp(width, 120, 2400);
            height = Math.Clamp(height, 80, 1600);

            (List<string> cats, List<ChartSeriesData> ser) = this.ExtractChartData(chart);
            Control? control = ChartFactory.Build(chart.Kind, cats, ser, width, height, chart.ShowDataLabels,
                chart.Title, chart.ShowLegend, chart.ValueAxisFormat, chart.AxisMinimum, chart.AxisMaximum, chart.AxisMajorStep);
            if (control is null)
            {
                continue;
            }

            // A small settings button (top-right) re-opens the builder for this chart.
            Button settings = new()
            {
                Content = "⚙",
                Width = 22,
                Height = 22,
                Padding = new Thickness(0),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Top,
                Margin = new Thickness(0, 3, 3, 0),
                Background = new SolidColorBrush(Color.FromArgb(0xCC, 0xFF, 0xFF, 0xFF)),
            };
            ToolTip.SetTip(settings, "Edit chart");
            SheetChart captured = chart;
            settings.Click += (_, _) => this.ChartEditRequested?.Invoke(captured);

            Grid content = new();
            content.Children.Add(control);
            content.Children.Add(settings);

            bool selected = ReferenceEquals(chart, this.selectedChart);
            Border card = new()
            {
                Background = Brushes.White,
                BorderBrush = selected ? SelectionBorder : this.gridLine,
                BorderThickness = new Thickness(selected ? 2 : 1),
                Child = content,
                Width = width,
                Height = height,
                Cursor = new Cursor(StandardCursorType.SizeAll),
            };
            Canvas.SetLeft(card, left);
            Canvas.SetTop(card, top);

            // Drag the card to reposition; the top-left corner snaps into a cell. A small movement
            // threshold keeps a plain click (or a click on the ⚙ button) from nudging the chart.
            double cardWidth = width;
            double cardHeight = height;
            bool pressed = false;
            bool moved = false;
            Point startPointer = default;
            double startLeft = left;
            double startTop = top;
            card.PointerPressed += (_, e) =>
            {
                if (IsWithin(e.Source as Visual, settings))
                {
                    return;
                }
                PointerPoint point = e.GetCurrentPoint(card);
                if (!point.Properties.IsLeftButtonPressed)
                {
                    return; // right-click etc.: let the ContextFlyout open (no rebuild)
                }
                // Left-click selects the chart (so the Delete key removes it) and focuses the grid.
                this.SelectChart(captured);
                pressed = true;
                moved = false;
                startPointer = e.GetPosition(this.cells);
                startLeft = Canvas.GetLeft(card);
                startTop = Canvas.GetTop(card);
                e.Pointer.Capture(card);
                this.chartDragging = true;
                e.Handled = true;
            };
            card.PointerMoved += (_, e) =>
            {
                if (!pressed || !ReferenceEquals(e.Pointer.Captured, card))
                {
                    return;
                }
                Point p = e.GetPosition(this.cells);
                double dx = p.X - startPointer.X;
                double dy = p.Y - startPointer.Y;
                if (!moved && Math.Abs(dx) < 4 && Math.Abs(dy) < 4)
                {
                    return;
                }
                moved = true;
                // Live snap: show the card landing flush in the nearest cell as it's dragged.
                int col = this.NearestColumn(Math.Max(0, startLeft + dx));
                int row = this.NearestRow(Math.Max(0, startTop + dy));
                Canvas.SetLeft(card, this.OffsetToColumn(col));
                Canvas.SetTop(card, this.OffsetToRow(row));
            };
            card.PointerReleased += (_, e) =>
            {
                if (!ReferenceEquals(e.Pointer.Captured, card))
                {
                    return;
                }
                e.Pointer.Capture(null);
                this.chartDragging = false;
                if (pressed && moved)
                {
                    int col = this.NearestColumn(Canvas.GetLeft(card));
                    int row = this.NearestRow(Canvas.GetTop(card));
                    this.sheet?.MoveChart(captured, row, col, cardWidth, cardHeight);
                }
                pressed = false;
                moved = false;
            };

            // Right-click menu with a Delete item (acts on this chart directly).
            MenuItem deleteItem = new() { Header = "Delete chart" };
            deleteItem.Click += (_, _) => this.sheet?.RemoveChart(captured);
            MenuFlyout menu = new();
            menu.Items.Add(deleteItem);
            card.ContextFlyout = menu;

            // Resize grip (bottom-right corner): grab and drag to resize like an image.
            Border grip = new()
            {
                Width = 14,
                Height = 14,
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 1, 1),
                Background = new SolidColorBrush(Color.FromArgb(0xCC, 0x1A, 0x73, 0xE8)),
                CornerRadius = new CornerRadius(2, 0, 2, 0),
                Cursor = new Cursor(StandardCursorType.BottomRightCorner),
            };
            ToolTip.SetTip(grip, "Drag to resize");
            content.Children.Add(grip);

            bool sizing = false;
            Point startSizePointer = default;
            double startW = width;
            double startH = height;
            grip.PointerPressed += (_, e) =>
            {
                sizing = true;
                startSizePointer = e.GetPosition(this.cells);
                startW = card.Width;
                startH = card.Height;
                e.Pointer.Capture(grip);
                this.chartDragging = true;   // reuse the guard: suppress row-adjust + block card drag
                e.Handled = true;
            };
            grip.PointerMoved += (_, e) =>
            {
                if (!sizing || !ReferenceEquals(e.Pointer.Captured, grip))
                {
                    return;
                }
                Point p = e.GetPosition(this.cells);
                double w = Math.Clamp(startW + (p.X - startSizePointer.X), 120, 2400);
                double h = Math.Clamp(startH + (p.Y - startSizePointer.Y), 80, 1600);
                card.Width = w;
                card.Height = h;
                control.Width = w;       // live-resize the chart itself
                control.Height = h;
            };
            grip.PointerReleased += (_, e) =>
            {
                if (!ReferenceEquals(e.Pointer.Captured, grip))
                {
                    return;
                }
                e.Pointer.Capture(null);
                this.chartDragging = false;
                if (sizing)
                {
                    this.sheet?.ResizeChart(captured, card.Width, card.Height);
                }
                sizing = false;
            };

            this.chartOverlay.Children.Add(card);
            this.chartCards.Add((card, control, settings, grip));
        }
    }

    // Renders each live chart card to a PNG snapshot with its grid position/size, for compositing
    // into an exported PDF. The on-card chrome (⚙ button + resize grip) is hidden during capture so
    // only the chart and its white card show. Must be called on the UI thread (charts are live
    // controls). Returns empty if the charts have not been rendered (e.g. headless use).
    public IReadOnlyList<ChartImage> RenderChartImages()
    {
        List<ChartImage> images = new();
        foreach ((Border card, Control chart, Control settings, Control grip) in this.chartCards)
        {
            double width = card.Width;
            double height = card.Height;
            if (double.IsNaN(width) || double.IsNaN(height) || width < 1 || height < 1)
            {
                continue;
            }

            double left = Canvas.GetLeft(card);
            double top = Canvas.GetTop(card);
            if (double.IsNaN(left))
            {
                left = 0;
            }

            if (double.IsNaN(top))
            {
                top = 0;
            }

            bool settingsVisible = settings.IsVisible;
            bool gripVisible = grip.IsVisible;
            settings.IsVisible = false;
            grip.IsVisible = false;
            try
            {
                PixelSize pixelSize = new((int)Math.Ceiling(width), (int)Math.Ceiling(height));
                using RenderTargetBitmap bitmap = new(pixelSize, new Vector(96, 96));
                bitmap.Render(card);
                using MemoryStream stream = new();
                bitmap.Save(stream, new PngBitmapEncoderOptions());
                images.Add(new ChartImage(stream.ToArray(), left, top, width, height));
            }
            catch
            {
                // Skip a chart that can't be snapshotted rather than failing the whole export.
            }
            finally
            {
                settings.IsVisible = settingsVisible;
                grip.IsVisible = gripVisible;
            }
        }

        return images;
    }

    // Resolves every chart on the sheet into the data the .xlsx chart writer needs: the range formulas
    // (qualified with the sheet name) plus the live values/labels so a saved workbook keeps its charts.
    public IReadOnlyList<ChartWriteData> BuildChartWriteData(string sheetName)
    {
        List<ChartWriteData> result = new();
        if (this.sheet is null)
        {
            return result;
        }

        foreach (SheetChart chart in this.sheet.Charts)
        {
            (List<string> categories, List<ChartSeriesData> series) = this.ExtractChartData(chart);

            List<ChartWriteSeries> written = new();
            for (int i = 0; i < chart.Series.Count; i++)
            {
                ChartSeriesSpec spec = chart.Series[i];
                IReadOnlyList<double> values = i < series.Count ? series[i].Values : Array.Empty<double>();
                written.Add(new ChartWriteSeries(spec.Title, RangesToFormula(sheetName, spec.Ranges), values));
            }

            // Reproduce the chart's on-screen rectangle exactly (same math as RefreshCharts): a
            // two-cell-anchored chart is sized by its from→to span, otherwise by its stored size. We
            // persist it as an anchor + explicit pixel extent so position and size survive the round-trip.
            double left = this.OffsetToColumn(chart.AnchorColumn) + chart.OffsetX;
            double top = this.OffsetToRow(chart.AnchorRow) + chart.OffsetY;
            double width = chart.Width;
            double height = chart.Height;
            if (chart.ToColumn is { } toColumn && chart.ToRow is { } toRow)
            {
                width = this.OffsetToColumn(toColumn) + chart.ToOffsetX - left;
                height = this.OffsetToRow(toRow) + chart.ToOffsetY - top;
            }

            width = Math.Clamp(width, 120, 2400);
            height = Math.Clamp(height, 80, 1600);

            string? categoryFormula = chart.Categories is { } range ? RangeToFormula(sheetName, range) : null;
            result.Add(new ChartWriteData(
                chart.Kind, chart.Title, chart.ShowLegend, chart.ShowDataLabels, chart.AxisMinimum,
                chart.AnchorRow, chart.AnchorColumn, chart.OffsetX, chart.OffsetY, width, height,
                written, categoryFormula, categories));
        }

        return result;
    }

    // Builds the OOXML value reference for a series: one sheet-qualified range, or a parenthesised
    // multi-area reference (e.g. ('Budget'!$C$7,'Budget'!$C$10)) when the series spans several ranges.
    private static string RangesToFormula(string sheetName, IReadOnlyList<CellRange> ranges)
    {
        List<string> parts = ranges
            .Where(static r => r.RowSpan > 0 && r.ColumnSpan > 0)
            .Select(r => RangeToFormula(sheetName, r))
            .ToList();
        return parts.Count > 1 ? $"({string.Join(",", parts)})" : parts.FirstOrDefault() ?? string.Empty;
    }

    // Builds an absolute, sheet-qualified A1 range formula (e.g. 'Budget'!$C$7:$C$10) for a chart part.
    private static string RangeToFormula(string sheetName, CellRange range)
    {
        string firstColumn = CellAddress.ColumnToLetters(range.Column);
        string lastColumn = CellAddress.ColumnToLetters(range.Column + Math.Max(1, range.ColumnSpan) - 1);
        int firstRow = range.Row + 1;
        int lastRow = range.Row + Math.Max(1, range.RowSpan);
        string a1 = range.RowSpan <= 1 && range.ColumnSpan <= 1
            ? $"${firstColumn}${firstRow}"
            : $"${firstColumn}${firstRow}:${lastColumn}${lastRow}";
        return $"'{sheetName.Replace("'", "''")}'!{a1}";
    }

    // Reads a chart's explicit series + category ranges into the data the factory needs, pulling the
    // live computed cell values.
    private (List<string> Categories, List<ChartSeriesData> Series) ExtractChartData(SheetChart chart)
    {
        string Display(CellAddress a) => this.sheet!.GetCell(a)?.DisplayText ?? string.Empty;

        // True when the range resolves to at least one non-empty cell in the loaded sheet, i.e. the
        // chart's data lives here and should be read live (so edits/recalcs update the chart). When
        // false (data on another sheet or outside the loaded range) we fall back to the file snapshot.
        bool RangeHasData(CellRange range)
        {
            if (range.RowSpan <= 0 || range.ColumnSpan <= 0)
            {
                return false;
            }

            return this.EnumerateRange(range).Any(a => !string.IsNullOrWhiteSpace(Display(a)));
        }

        List<ChartSeriesData> series = new();
        foreach (ChartSeriesSpec spec in chart.Series)
        {
            // A series can span several ranges/cells — concatenate their cells in order into one series.
            List<double> values = spec.Ranges.Any(RangeHasData)
                ? spec.Ranges.SelectMany(this.EnumerateRange).Select(a => ParseNumber(Display(a)) ?? double.NaN).ToList()
                : spec.CachedValues?.ToList() ?? new List<double>();
            series.Add(new ChartSeriesData(spec.Title, values));
        }

        int maxLen = series.Count == 0 ? 0 : series.Max(s => s.Values.Count);
        List<string> categories;
        if (chart.Categories is { } cats && RangeHasData(cats))
        {
            categories = this.EnumerateRange(cats).Select(Display).ToList();
        }
        else if (chart.CachedCategories is { Count: > 0 } cached)
        {
            categories = cached.ToList();
        }
        else if (chart.Categories is { } cats2)
        {
            categories = this.EnumerateRange(cats2).Select(Display).ToList();
        }
        else
        {
            categories = Enumerable.Range(1, maxLen).Select(i => i.ToString()).ToList();
        }

        // One bar/point per value: if the series has more values than category labels (e.g. the value
        // range grew), pad the labels with numbers so every value shows. Extra labels are trimmed.
        for (int i = categories.Count; i < maxLen; i++)
        {
            categories.Add((i + 1).ToString());
        }

        if (categories.Count > maxLen)
        {
            categories = categories.GetRange(0, maxLen);
        }

        return (categories, series);
    }

    // Yields a range's cell addresses in row-major order (clamped to the sheet).
    private IEnumerable<CellAddress> EnumerateRange(CellRange range)
    {
        int rowEnd = Math.Min(range.Row + range.RowSpan - 1, this.sheet!.RowCount - 1);
        int colEnd = Math.Min(range.Column + range.ColumnSpan - 1, this.sheet.ColumnCount - 1);
        for (int r = Math.Max(0, range.Row); r <= rowEnd; r++)
        {
            for (int c = Math.Max(0, range.Column); c <= colEnd; c++)
            {
                yield return new CellAddress(r, c);
            }
        }
    }

    // Tolerant numeric parse for cell display text ($, commas, %, parenthesised negatives).
    private static double? ParseNumber(string s)
    {
        if (string.IsNullOrWhiteSpace(s))
        {
            return null;
        }
        string t = s.Trim();
        bool negative = t.StartsWith('(') && t.EndsWith(')');
        if (negative)
        {
            t = t[1..^1];
        }
        bool percent = t.EndsWith('%');
        if (percent)
        {
            t = t[..^1];
        }
        t = t.Replace("$", "").Replace(",", "").Replace("€", "").Replace("£", "").Replace(" ", "").Trim();
        if (!double.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out double v))
        {
            return null;
        }
        if (percent)
        {
            v /= 100;
        }
        return negative ? -v : v;
    }

    private void BuildMergeMaps(int rows, int cols)
    {
        this.covered = new bool[rows, cols];
        this.spans = new (int, int)[rows, cols];
        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                this.spans[r, c] = (1, 1);
            }
        }

        foreach (MergeRange m in this.sheet!.Merges)
        {
            if (m.Row < 0 || m.Column < 0 || m.Row >= rows || m.Column >= cols)
            {
                continue;
            }
            this.spans[m.Row, m.Column] = (Math.Min(m.RowSpan, rows - m.Row), Math.Min(m.ColumnSpan, cols - m.Column));
            for (int r = m.Row; r < m.Row + m.RowSpan && r < rows; r++)
            {
                for (int c = m.Column; c < m.Column + m.ColumnSpan && c < cols; c++)
                {
                    if (r != m.Row || c != m.Column)
                    {
                        this.covered[r, c] = true;
                    }
                }
            }
        }
    }

    private Border HeaderCell(string text, int row, int column)
    {
        Border border = new()
        {
            Background = this.headerBackground,
            BorderBrush = this.gridLine,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Child = new TextBlock
            {
                Text = text,
                Foreground = this.headerText,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
            },
        };
        Grid.SetRow(border, row);
        Grid.SetColumn(border, column);
        return border;
    }

    // ----- Drag-to-resize columns/rows from the header boundaries (Excel-style) -----

    private const double MinColumnWidth = 24;
    private const double MinRowHeight = 16;
    private double dragStart;
    private double dragStartSize;
    private bool resizing; // suppress dynamic-row growth/trim while dragging a header divider

    // A thin grab strip on a column header's right edge: drag to resize that column.
    private Control ColumnResizeHandle(int column)
    {
        Border handle = new()
        {
            Background = Brushes.Transparent,
            Width = 6,
            HorizontalAlignment = HorizontalAlignment.Right,
            Margin = new Thickness(0, 0, -3, 0),
            Cursor = new Cursor(StandardCursorType.SizeWestEast),
        };
        Grid.SetColumn(handle, column);
        handle.PointerPressed += (_, e) =>
        {
            this.dragStart = e.GetPosition(this).X;
            this.dragStartSize = this.columnWidths[column];
            this.resizing = true;
            e.Pointer.Capture(handle);
            e.Handled = true;
        };
        handle.PointerMoved += (_, e) =>
        {
            if (!ReferenceEquals(e.Pointer.Captured, handle))
            {
                return;
            }
            double width = Math.Max(MinColumnWidth, this.dragStartSize + (e.GetPosition(this).X - this.dragStart));
            this.SetColumnWidth(column, width);
        };
        handle.PointerReleased += (_, e) =>
        {
            if (!ReferenceEquals(e.Pointer.Captured, handle))
            {
                return;
            }
            e.Pointer.Capture(null);
            this.resizing = false;
            this.Rebuild(); // resync overlays/selection/editor to the final size
        };
        return handle;
    }

    // A thin grab strip on a row header's bottom edge: drag to resize that row.
    private Control RowResizeHandle(int row)
    {
        Border handle = new()
        {
            Background = Brushes.Transparent,
            Height = 6,
            VerticalAlignment = VerticalAlignment.Bottom,
            Margin = new Thickness(0, 0, 0, -3),
            Cursor = new Cursor(StandardCursorType.SizeNorthSouth),
        };
        Grid.SetRow(handle, row);
        handle.PointerPressed += (_, e) =>
        {
            this.dragStart = e.GetPosition(this).Y;
            this.dragStartSize = this.rowHeights[row];
            this.resizing = true;
            e.Pointer.Capture(handle);
            e.Handled = true;
        };
        handle.PointerMoved += (_, e) =>
        {
            if (!ReferenceEquals(e.Pointer.Captured, handle))
            {
                return;
            }
            double height = Math.Max(MinRowHeight, this.dragStartSize + (e.GetPosition(this).Y - this.dragStart));
            this.SetRowHeight(row, height);
        };
        handle.PointerReleased += (_, e) =>
        {
            if (!ReferenceEquals(e.Pointer.Captured, handle))
            {
                return;
            }
            e.Pointer.Capture(null);
            this.resizing = false;
            this.Rebuild();
        };
        return handle;
    }

    // Live-updates a column's width across the header + cell grids and the model (no rebuild).
    private void SetColumnWidth(int column, double width)
    {
        if (column < 0 || column >= this.columnWidths.Length)
        {
            return;
        }
        this.columnWidths[column] = width;
        if (column < this.columnHeaders.ColumnDefinitions.Count)
        {
            this.columnHeaders.ColumnDefinitions[column].Width = new GridLength(width);
        }
        if (column < this.cells.ColumnDefinitions.Count)
        {
            this.cells.ColumnDefinitions[column].Width = new GridLength(width);
        }
        if (this.sheet is not null && column < this.sheet.ColumnCount)
        {
            this.sheet.Columns[column].Width = width;
        }
    }

    // Live-updates a row's height across the header + cell grids and the model (no rebuild).
    private void SetRowHeight(int row, double height)
    {
        if (row < 0 || row >= this.rowHeights.Length)
        {
            return;
        }
        this.rowHeights[row] = height;
        if (row < this.rowHeaders.RowDefinitions.Count)
        {
            this.rowHeaders.RowDefinitions[row].Height = new GridLength(height);
        }
        if (row < this.cells.RowDefinitions.Count)
        {
            this.cells.RowDefinitions[row].Height = new GridLength(height);
        }
        if (this.sheet is not null && row < this.sheet.RowCount)
        {
            this.sheet.Rows[row].Height = height;
        }
    }

    // How far a cell's text spills into adjacent EMPTY cells (Excel-style overflow): left and right
    // counts. Only plain text spills — numbers/currency stay put. Direction follows alignment
    // (left→right, right→left, center→both).
    private (int Left, int Right) ComputeSpill(int row, int column, CellViewModel cell)
    {
        if (this.sheet is null || cell.NumberFormat is not null)
        {
            return (0, 0); // number-formatted (e.g. currency) cells don't spill — Excel shows ####
        }
        // Spill when the DISPLAYED value is text (incl. a text-producing formula); numbers and formula
        // ERROR values (#DIV/0!, #NAME?, …) stay put — Excel never overflows those into empty cells.
        string display = cell.DisplayText;
        if (string.IsNullOrWhiteSpace(display) ||
            double.TryParse(display, NumberStyles.Any, CultureInfo.InvariantCulture, out _) ||
            IsErrorValue(display) ||
            this.FindDataBar(row, column) is not null)
        {
            return (0, 0);
        }

        TextAlignment align = cell.TextAlignment;
        int right = align != TextAlignment.Right ? this.EmptyRun(row, column, +1) : 0;
        int left = align != TextAlignment.Left ? this.EmptyRun(row, column, -1) : 0;
        return (left, right);
    }

    // An Excel error value (#DIV/0!, #VALUE!, #NAME?, #REF!, #NUM!, #N/A, #CIRC!, …).
    private static bool IsErrorValue(string display) =>
        display.Length > 1 && display[0] == '#' && (display[^1] is '!' or '?' || display == "#N/A");

    // Counts consecutive empty, non-merged, non-databar cells from (row,column) in a direction.
    private int EmptyRun(int row, int column, int direction)
    {
        int count = 0;
        for (int c = column + direction; c >= 0 && c < this.sheet!.ColumnCount && c < this.sheet.Rows[row].Cells.Count; c += direction)
        {
            if (this.covered![row, c] || this.FindDataBar(row, c) is not null)
            {
                break;
            }
            if (!string.IsNullOrWhiteSpace(this.sheet.Rows[row].Cells[c].DisplayText))
            {
                break;
            }
            count++;
        }
        return count;
    }

    // An overlay TextBlock that renders a spilling cell's full text across its empty neighbours. Not
    // hit-testable, so clicks fall through to the real 1×1 cells beneath (selection stays correct).
    private Control BuildSpillText(CellViewModel cell, int row, int startColumn, int span)
    {
        TextBlock text = new()
        {
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 2),
            TextTrimming = TextTrimming.CharacterEllipsis,
            FontSize = cell.FontSize,
            TextAlignment = cell.TextAlignment,
            IsHitTestVisible = false,
        };
        if (!string.IsNullOrWhiteSpace(cell.FontFamily))
        {
            text.FontFamily = new FontFamily(cell.FontFamily);
        }
        RenderInlines(text, cell);
        Grid.SetRow(text, row);
        Grid.SetColumn(text, startColumn);
        Grid.SetColumnSpan(text, span);
        return text;
    }

    private Control BuildCell(CellViewModel cell, int row, int column, int rowSpan, int columnSpan, bool spilling = false)
    {
        TextBlock text = new()
        {
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 2),
            // When the text spills into empty neighbours, an overlay TextBlock draws the full text on
            // top; the in-cell one then shouldn't show an ellipsis (the overlay covers it).
            TextTrimming = spilling ? TextTrimming.None : TextTrimming.CharacterEllipsis,
            FontSize = cell.FontSize,
            TextAlignment = cell.TextAlignment,
        };
        if (!string.IsNullOrWhiteSpace(cell.FontFamily))
        {
            text.FontFamily = new FontFamily(cell.FontFamily);
        }
        // A spilling cell's text is drawn entirely by the overlay; keep the in-cell one empty so the
        // two don't render on top of each other.
        if (!spilling)
        {
            RenderInlines(text, cell);
        }

        // Conditional-formatting data bar: a proportional bar behind (or instead of) the value.
        SheetDataBar? dataBar = this.FindDataBar(row, column);
        Grid? barGrid = null;
        if (dataBar is { } db)
        {
            barGrid = new Grid();
            Border fill = new() { Background = new SolidColorBrush(db.Color), Margin = new Thickness(1) };
            Grid.SetColumn(fill, 0);
            barGrid.Children.Add(fill);
            SetBarFraction(barGrid, db, cell);
            if (!db.ShowValue)
            {
                text.Inlines?.Clear(); // show only the bar (Excel "show bar only")
            }
        }

        Border border = new()
        {
            BorderBrush = this.gridLine,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Background = cell.Background ?? Brushes.Transparent,
            Child = barGrid is null ? text : Layered(barGrid, text),
        };
        Grid.SetRow(border, row);
        Grid.SetColumn(border, column);
        if (rowSpan > 1)
        {
            Grid.SetRowSpan(border, rowSpan);
        }
        if (columnSpan > 1)
        {
            Grid.SetColumnSpan(border, columnSpan);
        }

        CellAddress address = new(row, column);
        border.PointerPressed += (_, e) =>
        {
            // Right-click: select the cell first (unless it's already inside the selection), then let
            // the context flyout open over the right range.
            if (e.GetCurrentPoint(border).Properties.IsRightButtonPressed)
            {
                if (!this.InSelection(address))
                {
                    this.SelectOnly(address);
                }
                return;
            }
            if (e.ClickCount >= 2)
            {
                this.BeginEdit(address); // double-click edits
                return;
            }
            if (e.KeyModifiers.HasFlag(KeyModifiers.Shift))
            {
                this.ExtendSelectionTo(address); // shift+click extends from the anchor
                return;
            }
            // Single click selects this cell and starts a drag-select from it.
            this.SelectOnly(address);
            this.rangeDragging = true;
            e.Pointer.Capture(border);
        };
        border.PointerMoved += (_, e) =>
        {
            if (!this.rangeDragging || !ReferenceEquals(e.Pointer.Captured, border))
            {
                return;
            }
            Point p = e.GetPosition(this.cells);
            this.ExtendSelectionTo(new CellAddress(this.RowAtY(p.Y), this.ColAtX(p.X)));
        };
        border.PointerReleased += (_, e) =>
        {
            if (!ReferenceEquals(e.Pointer.Captured, border))
            {
                return;
            }
            e.Pointer.Capture(null);
            this.rangeDragging = false;
        };

        // Re-render live as the active cell is edited via the formula bar / toolbar. A font-size
        // change can alter the row height, so that triggers a full rebuild instead.
        PropertyChangedEventHandler handler = (_, e) =>
        {
            switch (e.PropertyName)
            {
                case nameof(CellViewModel.EffectiveRuns) or nameof(CellViewModel.Text):
                    // A spilling cell's text lives in a separate overlay whose span depends on the
                    // neighbours — rebuild so it (and the spill extent) recomputes cleanly.
                    if (spilling)
                    {
                        this.Rebuild();
                        break;
                    }
                    if (dataBar is { ShowValue: false })
                    {
                        text.Inlines?.Clear();
                    }
                    else
                    {
                        RenderInlines(text, cell);
                    }
                    if (barGrid is not null && dataBar is { } bar)
                    {
                        SetBarFraction(barGrid, bar, cell);
                    }
                    break;
                case nameof(CellViewModel.TextAlignment):
                    text.TextAlignment = cell.TextAlignment;
                    break;
                case nameof(CellViewModel.FontFamily):
                    text.FontFamily = string.IsNullOrWhiteSpace(cell.FontFamily) ? FontFamily.Default : new FontFamily(cell.FontFamily);
                    break;
                case nameof(CellViewModel.Background):
                    border.Background = cell.Background ?? Brushes.Transparent;
                    break;
                case nameof(CellViewModel.FontSize):
                    this.Rebuild();
                    break;
            }
        };
        cell.PropertyChanged += handler;
        this.cellSubscriptions.Add((cell, handler));
        return border;
    }

    private SheetDataBar? FindDataBar(int row, int column)
    {
        if (this.sheet is null)
        {
            return null;
        }
        foreach (SheetDataBar b in this.sheet.DataBars)
        {
            if (row >= b.Range.Row && row < b.Range.Row + b.Range.RowSpan &&
                column >= b.Range.Column && column < b.Range.Column + b.Range.ColumnSpan)
            {
                return b;
            }
        }
        return null;
    }

    // Sizes a data bar's fill (column 0) proportional to the cell value between min and max.
    private static void SetBarFraction(Grid barGrid, SheetDataBar bar, CellViewModel cell)
    {
        double value = ParseNumber(cell.DisplayText) ?? bar.Min;
        double fraction = bar.Max > bar.Min ? Math.Clamp((value - bar.Min) / (bar.Max - bar.Min), 0, 1) : 0;
        string f = fraction.ToString("0.####", CultureInfo.InvariantCulture);
        string rest = (1 - fraction).ToString("0.####", CultureInfo.InvariantCulture);
        barGrid.ColumnDefinitions = new ColumnDefinitions($"{f}*,{rest}*");
    }

    // Stacks the data bar behind the text within a cell.
    private static Control Layered(Control back, Control front)
    {
        Grid grid = new();
        grid.Children.Add(back);
        grid.Children.Add(front);
        return grid;
    }

    private static void RenderInlines(TextBlock target, CellViewModel cell)
    {
        target.Inlines ??= new InlineCollection();
        target.Inlines.Clear();
        foreach (CellRun r in cell.EffectiveRuns)
        {
            if (string.IsNullOrEmpty(r.Text))
            {
                continue;
            }
            Run run = new(r.Text)
            {
                FontWeight = r.Bold ? FontWeight.Bold : FontWeight.Normal,
                FontStyle = r.Italic ? FontStyle.Italic : FontStyle.Normal,
            };
            if (r.Underline || r.Strikethrough)
            {
                TextDecorationCollection decorations = new();
                if (r.Underline)
                {
                    decorations.Add(new TextDecoration { Location = TextDecorationLocation.Underline });
                }

                if (r.Strikethrough)
                {
                    decorations.Add(new TextDecoration { Location = TextDecorationLocation.Strikethrough });
                }

                run.TextDecorations = decorations;
            }

            if (r.Foreground is { } fg)
            {
                run.Foreground = new SolidColorBrush(fg);
            }
            target.Inlines.Add(run);
        }
    }

    private void HideSelectionVisuals()
    {
        this.selection.IsVisible = false;
        this.anchorBorder.IsVisible = false;
        this.editorHost.IsVisible = false;
        foreach (Border part in this.fillParts)
        {
            part.IsVisible = false;
        }
    }

    private static void SetFillPart(Border part, int row, int col, int rowSpan, int colSpan, bool show)
    {
        if (show && rowSpan > 0 && colSpan > 0)
        {
            Grid.SetRow(part, row);
            Grid.SetColumn(part, col);
            Grid.SetRowSpan(part, rowSpan);
            Grid.SetColumnSpan(part, colSpan);
            part.IsVisible = true;
        }
        else
        {
            part.IsVisible = false;
        }
    }

    // Moves the selection rectangle and the in-cell editor to the active cell (accounting for merge
    // spans).
    private void UpdateSelection()
    {
        if (this.sheet is null || this.spans is null || this.covered is null)
        {
            this.HideSelectionVisuals();
            return;
        }

        CellAddress addr = this.sheet.ActiveCell;
        if (addr.Row < 0 || addr.Column < 0 || addr.Row >= this.sheet.RowCount || addr.Column >= this.sheet.ColumnCount)
        {
            this.HideSelectionVisuals();
            return;
        }

        bool editing = !this.editorSuspended;

        // If the active cell is inside a merge, anchor the editor/selection on the merge's top-left.
        (int row, int col) = (addr.Row, addr.Column);
        if (this.covered[row, col])
        {
            foreach (MergeRange m in this.sheet.Merges)
            {
                if (row >= m.Row && row < m.Row + m.RowSpan && col >= m.Column && col < m.Column + m.ColumnSpan)
                {
                    (row, col) = (m.Row, m.Column);
                    break;
                }
            }
        }

        (int rowSpan, int colSpan) = this.spans[row, col];

        // The selection rectangle spans the whole anchor→focus range; a single cell uses the
        // merge-adjusted anchor bounds (no fill / anchor outline).
        int selRow, selCol, selRowSpan, selColSpan;
        if (this.rangeAnchor.Equals(this.rangeFocus))
        {
            (selRow, selCol, selRowSpan, selColSpan) = (row, col, rowSpan, colSpan);
            foreach (Border part in this.fillParts)
            {
                part.IsVisible = false;
            }
            this.anchorBorder.IsVisible = false;
        }
        else
        {
            selRow = Math.Min(this.rangeAnchor.Row, this.rangeFocus.Row);
            selCol = Math.Min(this.rangeAnchor.Column, this.rangeFocus.Column);
            selRowSpan = Math.Abs(this.rangeAnchor.Row - this.rangeFocus.Row) + 1;
            selColSpan = Math.Abs(this.rangeAnchor.Column - this.rangeFocus.Column) + 1;

            // Fill the range in 4 parts that tile around the anchor cell (so the anchor stays unfilled).
            (int ar1, int ac1, int ar2, int ac2) = (row, col, row + rowSpan - 1, col + colSpan - 1);
            int r2 = selRow + selRowSpan - 1;
            int c2 = selCol + selColSpan - 1;
            SetFillPart(this.fillParts[0], selRow, selCol, ar1 - selRow, selColSpan, ar1 > selRow);                 // above anchor
            SetFillPart(this.fillParts[1], ar2 + 1, selCol, r2 - ar2, selColSpan, ar2 < r2);                        // below anchor
            SetFillPart(this.fillParts[2], ar1, selCol, ar2 - ar1 + 1, ac1 - selCol, ac1 > selCol);                 // left of anchor
            SetFillPart(this.fillParts[3], ar1, ac2 + 1, ar2 - ar1 + 1, c2 - ac2, ac2 < c2);                        // right of anchor

            Grid.SetRow(this.anchorBorder, ar1);
            Grid.SetColumn(this.anchorBorder, ac1);
            Grid.SetRowSpan(this.anchorBorder, rowSpan);
            Grid.SetColumnSpan(this.anchorBorder, colSpan);
            this.anchorBorder.IsVisible = true;
        }
        Grid.SetRow(this.selection, selRow);
        Grid.SetColumn(this.selection, selCol);
        Grid.SetRowSpan(this.selection, selRowSpan);
        Grid.SetColumnSpan(this.selection, selColSpan);
        this.selection.IsVisible = true;

        // Position the shared editor over the anchor cell and load that cell's content.
        Grid.SetRow(this.editorHost, row);
        Grid.SetColumn(this.editorHost, col);
        Grid.SetRowSpan(this.editorHost, rowSpan);
        Grid.SetColumnSpan(this.editorHost, colSpan);
        CellViewModel? cell = this.sheet.GetCell(new CellAddress(row, col));
        // Opaque background (cell fill, else white) so the editor covers the display text beneath it.
        this.editorHost.Background = cell?.Background ?? this.surface;
        if (cell is not null && editing)
        {
            this.editor.Load(cell);
        }
        // Hidden while the formula bar holds the cursor, so the cell shows its read-only text instead.
        this.editorHost.IsVisible = editing;
    }

    // Hides the in-cell editor and hands the cursor to the formula bar (only one caret at a time).
    // The cell falls back to its read-only display text until a cell is clicked again.
    public void SuspendEditor()
    {
        this.editorSuspended = true;
        this.editorHost.IsVisible = false;
    }

    // Puts keyboard focus in the in-cell editor (shows the caret) so the user can type.
    public void FocusEditor(bool selectAll = false)
    {
        if (this.editorHost.IsVisible)
        {
            this.editor.FocusEditor(selectAll);
        }
    }

    // Brings the cursor into the active cell (un-suspending the in-cell editor) and focuses it.
    // selectAll highlights the whole text (used on double-click to edit).
    public void FocusCell(bool selectAll = false)
    {
        this.editorSuspended = false;
        this.UpdateSelection(); // shows + loads the editor
        // Focus on a later dispatcher pass: a freshly-shown Pro editor isn't realized yet this cycle,
        // so focusing it synchronously silently fails (the "double-click does nothing, triple works" bug).
        Dispatcher.UIThread.Post(() => this.FocusEditor(selectAll), DispatcherPriority.Background);
    }

    // ----- Toolbar formatting -----
    // When the cell is merely selected, format the WHOLE cell (shows immediately). When editing, format
    // the in-cell editor's selection (per-character). This is why Bold previously "only worked sometimes":
    // a selected-but-not-editing cell has no live editor selection to format.

    public void ToggleBold()
    {
        if (this.ReadOnly)
        {
            return;
        }
        if (!this.editorSuspended)
        {
            this.editor.ToggleBold();
            return;
        }
        this.sheet?.Checkpoint();
        List<CellViewModel> cells = this.SelectionCells();
        bool on = !cells.All(c => c.IsAllBold);
        foreach (CellViewModel c in cells)
        {
            c.SetWholeBold(on);
        }
    }

    public void ToggleItalic()
    {
        if (this.ReadOnly)
        {
            return;
        }
        if (!this.editorSuspended)
        {
            this.editor.ToggleItalic();
            return;
        }
        this.sheet?.Checkpoint();
        List<CellViewModel> cells = this.SelectionCells();
        bool on = !cells.All(c => c.IsAllItalic);
        foreach (CellViewModel c in cells)
        {
            c.SetWholeItalic(on);
        }
    }

    public void ToggleUnderline()
    {
        if (this.ReadOnly)
        {
            return;
        }
        if (!this.editorSuspended)
        {
            this.editor.ToggleUnderline();
            return;
        }
        this.sheet?.Checkpoint();
        List<CellViewModel> cells = this.SelectionCells();
        bool on = !cells.All(c => c.IsAllUnderline);
        foreach (CellViewModel c in cells)
        {
            c.SetWholeUnderline(on);
        }
    }

    public void ToggleStrikethrough()
    {
        if (this.ReadOnly)
        {
            return;
        }
        if (!this.editorSuspended)
        {
            this.editor.ToggleStrikethrough();
            return;
        }
        this.sheet?.Checkpoint();
        List<CellViewModel> cells = this.SelectionCells();
        bool on = !cells.All(c => c.IsAllStrikethrough);
        foreach (CellViewModel c in cells)
        {
            c.SetWholeStrikethrough(on);
        }
    }

    public void ApplyForegroundToCell(Color color)
    {
        if (this.ReadOnly)
        {
            return;
        }
        if (!this.editorSuspended)
        {
            this.editor.ApplyForeground(color);
            return;
        }
        this.sheet?.Checkpoint();
        foreach (CellViewModel c in this.SelectionCells())
        {
            c.SetWholeForeground(color);
        }
    }

    private CellViewModel? ActiveCell() => this.sheet?.GetCell(this.sheet.ActiveCell);

    // ----- Clipboard / range operations -----

    private MenuFlyout BuildCellMenu()
    {
        MenuFlyout menu = new();
        void Add(string header, Action act)
        {
            MenuItem item = new() { Header = header };
            item.Click += (_, _) => act();
            menu.Items.Add(item);
        }
        Add("Cut", this.CutSelection);
        Add("Copy", this.CopySelection);
        Add("Paste", this.PasteSelection);
        menu.Items.Add(new Separator());
        Add("Delete contents", this.ClearSelection);
        return menu;
    }

    private bool InSelection(CellAddress a)
    {
        CellRange r = this.SelectionRange;
        return a.Row >= r.Row && a.Row < r.Row + r.RowSpan && a.Column >= r.Column && a.Column < r.Column + r.ColumnSpan;
    }

    // Copies the selected range to the internal clipboard (rich) and the OS clipboard (TSV).
    public void CopySelection()
    {
        if (this.sheet is null)
        {
            return;
        }
        CellRange range = this.SelectionRange;
        CellData[][] block = new CellData[range.RowSpan][];
        StringBuilder tsv = new();
        for (int i = 0; i < range.RowSpan; i++)
        {
            block[i] = new CellData[range.ColumnSpan];
            for (int j = 0; j < range.ColumnSpan; j++)
            {
                CellViewModel? cell = this.sheet.GetCell(new CellAddress(range.Row + i, range.Column + j));
                block[i][j] = cell?.ToData() ?? CellData.Plain(string.Empty);
                if (j > 0)
                {
                    tsv.Append('\t');
                }
                tsv.Append(cell?.DisplayText ?? string.Empty);
            }
            tsv.Append('\n');
        }
        this.clipboard = block;
        this.clipboardText = tsv.ToString();
        this.SetOsClipboard(this.clipboardText);
    }

    public void CutSelection()
    {
        this.CopySelection();
        this.ClearSelection();
    }

    // Clears the selected cells' contents (keeps formatting).
    public void ClearSelection()
    {
        if (this.sheet is null || this.ReadOnly)
        {
            return;
        }
        this.sheet.Checkpoint();
        this.sheet.RunBatch(() =>
        {
            foreach (CellViewModel cell in this.SelectionCells())
            {
                cell.Text = string.Empty;
            }
        });
    }

    // Pastes the clipboard at the selection's top-left, growing the sheet as needed. Uses the rich
    // internal block unless the OS clipboard changed (e.g. copied from Excel), then TSV values.
    public async void PasteSelection()
    {
        if (this.sheet is null || this.ReadOnly)
        {
            return;
        }
        CellRange range = this.SelectionRange;
        string? osText = null;
        IClipboard? clip = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clip is not null)
        {
            try { osText = await clip.TryGetTextAsync(); } catch { }
        }

        CellData[][]? block = this.clipboard is not null && (osText is null || osText == this.clipboardText)
            ? this.clipboard
            : ParseTsv(osText);
        if (block is null || block.Length == 0)
        {
            return;
        }

        this.sheet.Checkpoint();
        this.sheet.RunBatch(() =>
        {
            for (int i = 0; i < block.Length; i++)
            {
                for (int j = 0; j < block[i].Length; j++)
                {
                    CellAddress addr = new(range.Row + i, range.Column + j);
                    this.sheet.EnsureSize(addr.Row + 1, addr.Column + 1);
                    this.sheet.GetCell(addr)?.ApplyData(block[i][j]);
                }
            }
        });
        this.Rebuild();
    }

    // Selects the whole sheet (Ctrl+A).
    public void SelectAll()
    {
        if (this.sheet is null || this.sheet.RowCount == 0 || this.sheet.ColumnCount == 0)
        {
            return;
        }
        this.editorSuspended = true;
        this.CellActivated?.Invoke(new CellAddress(0, 0));
        Dispatcher.UIThread.Post(() =>
        {
            this.rangeFocus = new CellAddress(this.sheet.RowCount - 1, this.sheet.ColumnCount - 1);
            this.UpdateSelection();
            this.Focus();
        });
    }

    private async void SetOsClipboard(string text)
    {
        IClipboard? clip = TopLevel.GetTopLevel(this)?.Clipboard;
        if (clip is not null)
        {
            try { await clip.SetTextAsync(text); } catch { }
        }
    }

    private static CellData[][]? ParseTsv(string? text)
    {
        if (string.IsNullOrEmpty(text))
        {
            return null;
        }
        List<string> lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n').ToList();
        if (lines.Count > 0 && lines[^1].Length == 0)
        {
            lines.RemoveAt(lines.Count - 1); // trailing newline
        }
        return lines.Select(line => line.Split('\t').Select(CellData.Plain).ToArray()).ToArray();
    }

    // Every cell in the current selection range (at least the active cell).
    public List<CellViewModel> SelectionCells()
    {
        List<CellViewModel> cells = new();
        if (this.sheet is null)
        {
            return cells;
        }
        CellRange range = this.SelectionRange;
        for (int r = range.Row; r < range.Row + range.RowSpan; r++)
        {
            for (int c = range.Column; c < range.Column + range.ColumnSpan; c++)
            {
                if (this.sheet.GetCell(new CellAddress(r, c)) is { } cell)
                {
                    cells.Add(cell);
                }
            }
        }
        if (cells.Count == 0 && this.ActiveCell() is { } active)
        {
            cells.Add(active);
        }
        return cells;
    }

    // Current B/I/U state to reflect in the toolbar: the editor selection while editing, else the
    // active cell's whole-cell flags.
    public (bool Bold, bool Italic, bool Underline, bool Strikethrough) FormatState()
    {
        if (!this.editorSuspended)
        {
            return this.editor.GetSelectionFlags();
        }
        CellViewModel? c = this.ActiveCell();
        return (c?.Bold ?? false, c?.Italic ?? false, c?.Underline ?? false, c?.Strikethrough ?? false);
    }

    // Raised when the editing selection changes, so the host can re-sync the toolbar.
    public event EventHandler? SelectionFormatChanged;

    // Raised when a chart's settings button is clicked (host opens the editor).
    public event Action<SheetChart>? ChartEditRequested;

    public bool RichEditingAvailable => this.editor.IsRich;
}
