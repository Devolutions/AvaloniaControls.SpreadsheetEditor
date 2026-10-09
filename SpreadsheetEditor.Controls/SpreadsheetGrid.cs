namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Primitives;
using Avalonia.Controls.Shapes;
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
using global::SpreadsheetEditor.Controls.Localization;
using global::SpreadsheetEditor.Controls.Printing;
using global::SpreadsheetEditor.Controls.ViewModels;
using S = global::SpreadsheetEditor.Controls.Localization.EditorString;
using static global::SpreadsheetEditor.Controls.Localization.EditorText;

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

    // Excel-style dark mode: the CHROME (row/column headers, and the host's toolbar/formula bar/tabs)
    // follows the light/dark theme, but the SHEET PANEL does not — cells, gridlines, cell text and
    // charts stay light in both themes, exactly like Excel, which keeps the sheet a white page.
    // These are fixed rather than theme-resolved for that reason.
    private static readonly IBrush SheetGridLine = new SolidColorBrush(Color.FromArgb(0x30, 0, 0, 0));
    private static readonly IBrush SheetSurface = Brushes.White;

    // Cell text is set explicitly because the sheet stays light: the inherited dark-theme foreground
    // would be near-white and invisible against the white cells. Per-run colors still override it.
    private static readonly IBrush CellText = new SolidColorBrush(Color.FromRgb(0x20, 0x20, 0x20));
    private static readonly IBrush LinkText = new SolidColorBrush(Color.FromRgb(0x05, 0x63, 0xC1));

    // Chrome brushes, re-resolved from the active ThemeVariant in ApplyTheme() (on attach + on theme
    // change) so the headers darken while the sheet panel stays light.
    private IBrush headerBackground = new SolidColorBrush(Color.FromRgb(0xF3, 0xF3, 0xF3));
    private IBrush headerText = new SolidColorBrush(Color.FromRgb(0x60, 0x60, 0x60));
    private IBrush headerLine = new SolidColorBrush(Color.FromArgb(0x30, 0, 0, 0));

    private readonly ScrollViewer cellScroller = new()
    { HorizontalScrollBarVisibility = ScrollBarVisibility.Auto, VerticalScrollBarVisibility = ScrollBarVisibility.Auto };

    private readonly ScrollViewer columnHeaderScroller = new()
    { HorizontalScrollBarVisibility = ScrollBarVisibility.Hidden, VerticalScrollBarVisibility = ScrollBarVisibility.Disabled };

    private readonly ScrollViewer rowHeaderScroller = new()
    { HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled, VerticalScrollBarVisibility = ScrollBarVisibility.Hidden };

    private readonly Canvas columnHeaders = new();
    private readonly Canvas rowHeaders = new();
    private readonly Canvas cells = new();
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
    private bool editorSuspendedField = true;

    // Tracks an active editing session so that when it ends (commit) we recompute text spill once —
    // a cell whose text just grew past its width needs to overflow into empty neighbours (Excel-style),
    // which the spill overlay only picks up on a rebuild. Done on commit, not per keystroke.
    private bool editingSession;
    private CellAddress? editingAddress;
    private bool applyingCellFormat;

    private bool editorSuspended
    {
        get => this.editorSuspendedField;
        set
        {
            if (!value)
            {
                this.editorSuspendedField = false;
                if (!this.editingSession)
                {
                    this.editingAddress = this.sheet?.ActiveCell;
                    this.editingSession = true; // editing started
                }
                return;
            }

            this.editorSuspendedField = true;
            if (this.editingSession)
            {
                this.editingSession = false;
                CellAddress? address = this.editingAddress;
                this.editingAddress = null;
                if (this.sheet is { } editedSheet)
                {
                    // Text changes were intentionally skipped while the opaque editor covered the
                    // display cell. Refresh only that cell now; its subscriber requests a rebuild
                    // only if text overflow actually requires a different layout.
                    if (address is { } editedAddress)
                    {
                        editedSheet.GetCell(editedAddress)?.RefreshDisplay();
                    }
                    this.EditCompleted?.Invoke(editedSheet);
                }
            }
        }
    }

    private bool adjustingRows;            // reentrancy guard for dynamic row growth/trim
    private bool adjustingColumns;         // reentrancy guard for dynamic column growth/trim
    private const int RowBuffer = 25;       // rows kept past the last visible/used row
    private const int ColumnBuffer = 8;     // columns kept past the last visible/used column

    // Only cell/header controls near the viewport are materialized. Canvas extents preserve the full
    // scrollable sheet without allocating a RowDefinition and ColumnDefinition for every row/cell.
    private const int ViewportRowBuffer = 8;
    private const int ViewportColumnBuffer = 2;
    private bool viewportRenderPending;
    private int viewportRenderGeneration;
    private (int FirstRow, int LastRow, int FirstColumn, int LastColumn)? renderedViewport;

    private SpreadsheetViewModel? sheet;

    // Zoom multiplier applied to column widths, row heights, header sizes, and font sizes (1.0 = 100%).
    // Stored column widths / row heights are the LOGICAL sizes × zoom, so all coordinate math and the
    // pointer→cell mapping stay consistent; the model keeps the unscaled logical sizes.
    private double zoom = 1.0;
    private const double MinZoom = 0.5;
    private const double MaxZoom = 3.0;
    private const double ZoomStep = 0.1;

    private double[] columnWidths = Array.Empty<double>();
    private double[] rowHeights = Array.Empty<double>();
    private double[] columnOffsets = Array.Empty<double>();
    private double[] rowOffsets = Array.Empty<double>();
    private bool[,]? covered;
    private (int RowSpan, int ColumnSpan)[,]? spans;
    private readonly List<(CellViewModel Cell, PropertyChangedEventHandler Handler)> cellSubscriptions = new();
    private readonly Dictionary<Color, IBrush> gradientBarBrushes = new();

    public event Action<CellAddress>? CellActivated;

    // The host owns worksheet tabs; it resolves links that target a different sheet.
    public event Action<string, CellAddress>? WorksheetLinkRequested;

    // Raised once when an in-cell editing session is committed by Enter/Tab, selecting another cell,
    // moving focus to the formula bar, or otherwise returning to selection mode.
    public event Action<SpreadsheetViewModel>? EditCompleted;

    public bool IsEditing => !this.editorSuspended;

    // When true, the grid blocks all content mutation (editing, typing, paste/cut/clear, formatting).
    public bool ReadOnly { get; set; }

    // Returns one viewport plus a small scroll buffer for the lightweight worksheet preview. The
    // fallback covers startup before Avalonia has completed the first layout pass.
    internal (int Rows, int Columns) SuggestedPreviewSize()
    {
        double height = this.cellScroller.Viewport.Height;
        double width = this.cellScroller.Viewport.Width;
        if (!double.IsFinite(height) || height <= 0)
        {
            height = Math.Max(480, this.Bounds.Height - ColumnHeaderHeight);
        }
        if (!double.IsFinite(width) || width <= 0)
        {
            width = Math.Max(800, this.Bounds.Width - RowHeaderWidth);
        }

        int rows = (int)Math.Ceiling(height / Math.Max(1, DefaultRowHeight * this.zoom)) + (ViewportRowBuffer * 2);
        int columns = (int)Math.Ceiling(width / Math.Max(1, 96 * this.zoom)) + (ViewportColumnBuffer * 2);
        return (Math.Clamp(rows, 32, 128), Math.Clamp(columns, 12, 64));
    }

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
            BorderBrush = this.headerLine,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Width = RowHeaderWidth,
            Height = ColumnHeaderHeight,
        };
        Grid.SetRow(this.cornerBox, 0); Grid.SetColumn(this.cornerBox, 0);

        this.columnHeaderScroller.Content = this.columnHeaders;
        Grid.SetRow(this.columnHeaderScroller, 0); Grid.SetColumn(this.columnHeaderScroller, 1);

        this.rowHeaderScroller.Content = this.rowHeaders;
        Grid.SetRow(this.rowHeaderScroller, 1); Grid.SetColumn(this.rowHeaderScroller, 0);

        this.imageOverlay.ZIndex = 3;
        this.chartOverlay.ZIndex = 4;
        foreach (Border part in this.fillParts)
        {
            part.ZIndex = 5;
        }
        this.selection.ZIndex = 6;
        this.anchorBorder.ZIndex = 7;
        this.editorHost.ZIndex = 8;
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
            this.ScheduleViewportRender();
        };

        // Ctrl + mouse wheel zooms (like Excel). Handled in the tunnel so the ScrollViewer doesn't
        // also scroll on the same gesture.
        this.AddHandler(PointerWheelChangedEvent, (_, e) =>
        {
            if (e.KeyModifiers.HasFlag(KeyModifiers.Control))
            {
                if (e.Delta.Y > 0)
                {
                    this.ZoomIn();
                }
                else if (e.Delta.Y < 0)
                {
                    this.ZoomOut();
                }

                e.Handled = true;
            }
        }, RoutingStrategies.Tunnel);

        // Right-click menu (acts on the selection).
        this.cells.ContextFlyout = this.BuildCellMenu();

        root.Children.Add(this.cornerBox);
        root.Children.Add(this.columnHeaderScroller);
        root.Children.Add(this.rowHeaderScroller);
        root.Children.Add(this.cellScroller);
        this.Content = root;

        // Crisper text and edges (esp. for small, zoomed-out cells) and high-quality scaling for any
        // bitmap content (chart snapshots).
        TextOptions.SetTextRenderingMode(this, TextRenderingMode.Antialias);
        RenderOptions.SetEdgeMode(this, EdgeMode.Antialias);
        RenderOptions.SetBitmapInterpolationMode(this, BitmapInterpolationMode.HighQuality);
        this.UseLayoutRounding = true;

        // Intercept navigation keys before the in-cell editor (tunnel) so arrows/Enter/Tab move the
        // selection. Typing on a selected cell starts editing. The grid is focusable so it can receive
        // these keys while in selection mode (no editor focused).
        this.Focusable = true;
        this.AddHandler(KeyDownEvent, this.OnNavigationKey, RoutingStrategies.Tunnel);
        this.AddHandler(TextInputEvent, this.OnTextInput, RoutingStrategies.Tunnel);

        this.cells.Background = SheetSurface;
        this.cellScroller.Background = SheetSurface;
        // Re-resolve the palette whenever the light/dark theme changes at runtime.
        this.ActualThemeVariantChanged += (_, _) => this.ApplyTheme();
    }

    protected override void OnAttachedToVisualTree(VisualTreeAttachmentEventArgs e)
    {
        base.OnAttachedToVisualTree(e);
        this.ApplyTheme(); // ActualThemeVariant is resolved once attached to the tree
    }

    // Resolves the CHROME brushes (row/column headers) from the active light/dark theme and re-renders.
    // The sheet panel — surface, gridlines, cell text, charts — is deliberately theme-independent: Excel's
    // dark mode darkens the surrounding UI but keeps the cells a light page. See the brush declarations.
    private void ApplyTheme()
    {
        bool dark = this.ActualThemeVariant == ThemeVariant.Dark;
        this.headerBackground = new SolidColorBrush(dark ? Color.FromRgb(0x2D, 0x2D, 0x30) : Color.FromRgb(0xF3, 0xF3, 0xF3));
        this.headerText = new SolidColorBrush(dark ? Color.FromRgb(0xC8, 0xC8, 0xC8) : Color.FromRgb(0x60, 0x60, 0x60));
        this.headerLine = new SolidColorBrush(dark ? Color.FromArgb(0x28, 0xFF, 0xFF, 0xFF) : Color.FromArgb(0x30, 0, 0, 0));

        this.cornerBox.Background = this.headerBackground;
        this.cornerBox.BorderBrush = this.headerLine;
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
                    case Key.OemPlus or Key.Add: this.ZoomIn(); e.Handled = true; return;
                    case Key.OemMinus or Key.Subtract: this.ZoomOut(); e.Handled = true; return;
                    case Key.D0 or Key.NumPad0: this.ResetZoom(); e.Handled = true; return;
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
        return IndexAtOffset(this.columnOffsets, x);
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
        // The editor has only just been made visible. Focusing at background priority lets its rich
        // text surface finish realization first; otherwise the grid keeps focus and only this initial
        // character is accepted.
        Dispatcher.UIThread.Post(() => this.FocusEditor(), DispatcherPriority.Background);
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

    // Selects and scrolls to a cell programmatically (used by find/replace and host integrations).
    public void SelectCell(CellAddress address)
    {
        if (this.sheet is null || !this.sheet.IsValid(address))
        {
            return;
        }

        this.MoveSelection(address);
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

    // Current zoom level (1.0 = 100%). Setting it re-lays-out the grid at the new scale.
    public double Zoom
    {
        get => this.zoom;
        set
        {
            double clamped = Math.Clamp(value, MinZoom, MaxZoom);
            if (Math.Abs(clamped - this.zoom) < 0.0001)
            {
                return;
            }

            this.zoom = clamped;
            if (this.sheet is not null)
            {
                this.Rebuild();
                // Zooming out shrinks every cell, so the viewport now spans more rows/columns than the
                // grid was built for — grow to refill it (once layout has the new, smaller extent).
                Dispatcher.UIThread.Post(
                    () =>
                    {
                        this.AdjustColumnsForScroll();
                        this.AdjustRowsForScroll();
                    },
                    DispatcherPriority.Background);
            }

            this.ZoomChanged?.Invoke(this.zoom);
        }
    }

    // Raised when the zoom level changes (so a host can show a "100%" indicator).
    public event Action<double>? ZoomChanged;

    public void ZoomIn() => this.Zoom = this.zoom + ZoomStep;

    public void ZoomOut() => this.Zoom = this.zoom - ZoomStep;

    public void ResetZoom() => this.Zoom = 1.0;

    // Rebuilds the entire grid from the current sheet (call after structural changes / load).
    public void Rebuild()
    {
        this.renderedViewport = null;

        if (this.sheet is null || this.sheet.RowCount == 0 || this.sheet.ColumnCount == 0)
        {
            this.ClearCellSubscriptions();
            this.columnHeaders.Children.Clear();
            this.rowHeaders.Children.Clear();
            this.cells.Children.Clear();
            this.imageOverlay.Children.Clear();
            this.chartOverlay.Children.Clear();
            this.columnHeaders.Width = 0;
            this.rowHeaders.Height = 0;
            this.cells.Width = 0;
            this.cells.Height = 0;
            return;
        }

        int cols = this.sheet.ColumnCount;
        int rows = this.sheet.RowCount;

        // Column widths / row heights are stored as logical size × zoom so layout and pointer mapping
        // stay consistent at any scale (the model keeps the unscaled logical sizes).
        this.columnWidths = new double[cols];
        this.columnOffsets = new double[cols + 1];
        for (int c = 0; c < cols; c++)
        {
            double logical = this.sheet.Columns[c].Width > 0 ? this.sheet.Columns[c].Width : this.sheet.DefaultColumnWidth;
            this.columnWidths[c] = logical * this.zoom;
            this.columnOffsets[c + 1] = this.columnOffsets[c] + this.columnWidths[c];
        }

        this.rowHeights = new double[rows];
        this.rowOffsets = new double[rows + 1];
        for (int r = 0; r < rows; r++)
        {
            if (this.sheet.Rows[r].IsHidden)
            {
                this.rowHeights[r] = 0;
            }
            else if (this.sheet.Rows[r].Height is { } explicitHeight)
            {
                this.rowHeights[r] = explicitHeight * this.zoom;
            }
            else
            {
                double maxFont = this.sheet.Rows[r].Cells.Count == 0 ? CellViewModel.DefaultFontSize : this.sheet.Rows[r].Cells.Max(c => c.FontSize);
                this.rowHeights[r] = Math.Max(DefaultRowHeight, maxFont + 8) * this.zoom;
            }
            this.rowOffsets[r + 1] = this.rowOffsets[r] + this.rowHeights[r];
        }

        this.BuildMergeMaps(rows, cols);

        this.UpdateCanvasExtents();

        // The frozen top-left corner must match the (zoom-scaled) header sizes, or a gap appears below
        // the column headers / beside the row headers at any zoom other than 100%.
        this.cornerBox.Width = RowHeaderWidth * this.zoom;
        this.cornerBox.Height = ColumnHeaderHeight * this.zoom;

        this.imageOverlay.Children.Clear();
        this.BuildImages(rows, cols);
        this.RefreshCharts();
        this.RenderViewport(force: true);

        // The first rebuild can happen before layout gives the ScrollViewer a viewport. Refresh once
        // after layout so the initial materialized window exactly fits the visible area.
        Dispatcher.UIThread.Post(this.ScheduleViewportRender, DispatcherPriority.Background);
    }

    private void UpdateCanvasExtents()
    {
        double width = this.columnOffsets.Length == 0 ? 0 : this.columnOffsets[^1];
        double height = this.rowOffsets.Length == 0 ? 0 : this.rowOffsets[^1];
        double headerHeight = ColumnHeaderHeight * this.zoom;
        double headerWidth = RowHeaderWidth * this.zoom;

        this.columnHeaders.Width = width;
        this.columnHeaders.Height = headerHeight;
        this.rowHeaders.Width = headerWidth;
        this.rowHeaders.Height = height;
        this.cells.Width = width;
        this.cells.Height = height;
        this.imageOverlay.Width = width;
        this.imageOverlay.Height = height;
        this.chartOverlay.Width = width;
        this.chartOverlay.Height = height;
        Canvas.SetLeft(this.imageOverlay, 0);
        Canvas.SetTop(this.imageOverlay, 0);
        Canvas.SetLeft(this.chartOverlay, 0);
        Canvas.SetTop(this.chartOverlay, 0);
    }

    private void ClearCellSubscriptions()
    {
        foreach ((CellViewModel cell, PropertyChangedEventHandler handler) in this.cellSubscriptions)
        {
            cell.PropertyChanged -= handler;
        }
        this.cellSubscriptions.Clear();
    }

    private void ScheduleViewportRender()
    {
        if (this.viewportRenderPending)
        {
            return;
        }

        this.viewportRenderPending = true;
        Dispatcher.UIThread.Post(() =>
        {
            this.viewportRenderPending = false;
            this.RenderViewport();
        }, DispatcherPriority.Render);
    }

    private (int FirstRow, int LastRow, int FirstColumn, int LastColumn) VisibleViewport()
    {
        int rows = this.rowHeights.Length;
        int columns = this.columnWidths.Length;
        double height = this.cellScroller.Viewport.Height;
        double width = this.cellScroller.Viewport.Width;
        Vector offset = this.cellScroller.Offset;

        int firstRow = this.RowAtY(Math.Max(0, offset.Y));
        int firstColumn = this.ColAtX(Math.Max(0, offset.X));
        int lastRow = height > 0
            ? this.RowAtY(offset.Y + height)
            : Math.Min(rows - 1, firstRow + 50);
        int lastColumn = width > 0
            ? this.ColAtX(offset.X + width)
            : Math.Min(columns - 1, firstColumn + 15);

        return (firstRow, lastRow, firstColumn, lastColumn);
    }

    // Replaces only the visible cell/header controls. Cached sheet-wide offsets position them on
    // canvases whose dimensions provide the full scroll extent.
    private void RenderViewport(bool force = false)
    {
        if (this.sheet is null || this.rowHeights.Length == 0 || this.columnWidths.Length == 0 ||
            this.covered is null || this.spans is null)
        {
            return;
        }

        (int FirstRow, int LastRow, int FirstColumn, int LastColumn) visible = this.VisibleViewport();
        if (!force && this.renderedViewport is { } current)
        {
            const int rowGuard = 2;
            const int columnGuard = 1;
            bool rowsCovered = visible.FirstRow >= current.FirstRow && visible.LastRow <= current.LastRow &&
                               (current.FirstRow == 0 || visible.FirstRow - current.FirstRow > rowGuard) &&
                               (current.LastRow == this.rowHeights.Length - 1 || current.LastRow - visible.LastRow > rowGuard);
            bool columnsCovered = visible.FirstColumn >= current.FirstColumn && visible.LastColumn <= current.LastColumn &&
                                  (current.FirstColumn == 0 || visible.FirstColumn - current.FirstColumn > columnGuard) &&
                                  (current.LastColumn == this.columnWidths.Length - 1 || current.LastColumn - visible.LastColumn > columnGuard);
            if (rowsCovered && columnsCovered)
            {
                return;
            }
        }

        (int FirstRow, int LastRow, int FirstColumn, int LastColumn) viewport = (
            Math.Max(0, visible.FirstRow - ViewportRowBuffer),
            Math.Min(this.rowHeights.Length - 1, visible.LastRow + ViewportRowBuffer),
            Math.Max(0, visible.FirstColumn - ViewportColumnBuffer),
            Math.Min(this.columnWidths.Length - 1, visible.LastColumn + ViewportColumnBuffer));
        this.renderedViewport = viewport;

        this.ClearCellSubscriptions();
        this.columnHeaders.Children.Clear();
        this.rowHeaders.Children.Clear();
        this.cells.Children.Clear();
        int generation = ++this.viewportRenderGeneration;

        for (int column = viewport.FirstColumn; column <= viewport.LastColumn; column++)
        {
            Border header = this.HeaderCell(CellAddress.ColumnToLetters(column));
            int selectedColumn = column;
            header.ContextFlyout = this.BuildColumnMenu(selectedColumn);
            header.PointerPressed += (_, e) =>
            {
                if (e.GetCurrentPoint(header).Properties.IsRightButtonPressed)
                {
                    this.SelectColumn(selectedColumn);
                }
            };
            this.PlaceColumnHeader(header, column);
            this.columnHeaders.Children.Add(header);
            this.columnHeaders.Children.Add(this.ColumnResizeHandle(column));
        }
        // Persistent layers use explicit Z-order so visible cells can be created in short batches
        // without ever covering images, charts, selection, or the editor.
        this.cells.Children.Add(this.imageOverlay);
        this.cells.Children.Add(this.chartOverlay);
        foreach (Border part in this.fillParts)
        {
            this.cells.Children.Add(part);
        }
        this.cells.Children.Add(this.selection);
        this.cells.Children.Add(this.anchorBorder);
        this.cells.Children.Add(this.editorHost);
        this.UpdateSelection();

        HashSet<(int Row, int Column)> renderedCells = new();
        bool renderedCrossingMerges = false;
        void RenderCrossingMerges()
        {
            if (this.sheet is null || this.spans is null)
            {
                return;
            }

            // A merge can extend into the visible area even when its anchor is above or left of
            // the materialized window. Draw it before yielding to the off-screen buffer batches.
            foreach (MergeRange merge in this.sheet.Merges)
            {
                int mergeLastRow = merge.Row + merge.RowSpan - 1;
                int mergeLastColumn = merge.Column + merge.ColumnSpan - 1;
                bool intersects = merge.Row <= viewport.LastRow && mergeLastRow >= viewport.FirstRow &&
                                  merge.Column <= viewport.LastColumn && mergeLastColumn >= viewport.FirstColumn;
                if (!intersects || renderedCells.Contains((merge.Row, merge.Column)) ||
                    merge.Row < 0 || merge.Column < 0 || merge.Row >= this.sheet.RowCount ||
                    merge.Column >= this.sheet.Rows[merge.Row].Cells.Count ||
                    this.sheet.Rows[merge.Row].IsHidden)
                {
                    continue;
                }

                (int rowSpan, int columnSpan) = this.spans[merge.Row, merge.Column];
                this.cells.Children.Add(this.BuildCell(
                    this.sheet.Rows[merge.Row].Cells[merge.Column], merge.Row, merge.Column, rowSpan, columnSpan));
                renderedCells.Add((merge.Row, merge.Column));
            }
        }
        int nextRow = viewport.FirstRow;
        Action? renderBatch = null;
        renderBatch = () =>
        {
            if (generation != this.viewportRenderGeneration || this.sheet is null ||
                this.covered is null || this.spans is null)
            {
                return;
            }

            // Paint every *visible* row in this dispatcher turn. Only the off-screen buffer is
            // incremental; yielding while the visible area is still empty causes a noticeable
            // top-to-bottom refill on opening a sheet or rebuilding after an edit.
            bool paintingVisible = nextRow <= visible.LastRow;
            int batchLastRow = paintingVisible ? visible.LastRow : viewport.LastRow;
            int rowBudget = paintingVisible ? int.MaxValue : 4;
            int visibleRows = 0;
            for (; nextRow <= batchLastRow && visibleRows < rowBudget; nextRow++)
            {
                int row = nextRow;
                if (this.sheet.Rows[row].IsHidden)
                {
                    continue;
                }
                visibleRows++;
                Border header = this.HeaderCell((row + 1).ToString());
                int selectedRow = row;
                header.ContextFlyout = this.BuildRowMenu(selectedRow);
                header.PointerPressed += (_, e) =>
                {
                    if (e.GetCurrentPoint(header).Properties.IsRightButtonPressed)
                    {
                        this.SelectRow(selectedRow);
                    }
                };
                this.PlaceRowHeader(header, row);
                this.rowHeaders.Children.Add(header);
                this.rowHeaders.Children.Add(this.RowResizeHandle(row));

                Dictionary<int, (int Left, int Right)> spillMap = new();
                HashSet<int> noRightBorder = new();
                int lastCellColumn = Math.Min(viewport.LastColumn, this.sheet.Rows[row].Cells.Count - 1);
                for (int column = viewport.FirstColumn; column <= lastCellColumn; column++)
                {
                    if (this.covered[row, column] || this.spans[row, column].ColumnSpan != 1)
                    {
                        continue;
                    }

                    (int spillLeft, int spillRight) = this.ComputeSpill(row, column, this.sheet.Rows[row].Cells[column]);
                    if (spillLeft + spillRight == 0)
                    {
                        continue;
                    }

                    spillMap[column] = (spillLeft, spillRight);
                    for (int hiddenBorder = column - spillLeft; hiddenBorder < column + spillRight; hiddenBorder++)
                    {
                        noRightBorder.Add(hiddenBorder);
                    }
                }

                for (int column = viewport.FirstColumn; column <= lastCellColumn; column++)
                {
                    if (this.covered[row, column] || renderedCells.Contains((row, column)))
                    {
                        continue;
                    }

                    (int rowSpan, int columnSpan) = this.spans[row, column];
                    CellViewModel cell = this.sheet.Rows[row].Cells[column];
                    bool spilling = spillMap.ContainsKey(column);
                    this.cells.Children.Add(this.BuildCell(
                        cell, row, column, rowSpan, columnSpan, spilling,
                        omitRightBorder: noRightBorder.Contains(column)));
                    renderedCells.Add((row, column));

                    if (spilling)
                    {
                        (int spillLeft, int spillRight) = spillMap[column];
                        Control overlay = this.BuildSpillText(cell, row, column - spillLeft, spillLeft + 1 + spillRight);
                        overlay.ZIndex = 2;
                        this.cells.Children.Add(overlay);
                    }
                }
            }

            if (!renderedCrossingMerges && nextRow > visible.LastRow)
            {
                renderedCrossingMerges = true;
                RenderCrossingMerges();
            }

            if (nextRow <= viewport.LastRow)
            {
                // A short timer guarantees the dispatcher can process input and paint between slices;
                // reposting immediately at the same priority can be drained in one long render pass.
                DispatcherTimer.RunOnce(renderBatch!, TimeSpan.FromMilliseconds(1), DispatcherPriority.Render);
                return;
            }

        };

        renderBatch();
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
        return this.columnOffsets.Length == 0 ? 0 : this.columnOffsets[Math.Clamp(column, 0, this.columnOffsets.Length - 1)];
    }

    private double OffsetToRow(int row)
    {
        return this.rowOffsets.Length == 0 ? 0 : this.rowOffsets[Math.Clamp(row, 0, this.rowOffsets.Length - 1)];
    }

    // The row index at a given Y pixel within the cell area.
    private int RowAtY(double y) => IndexAtOffset(this.rowOffsets, y);

    private static int IndexAtOffset(double[] offsets, double value)
    {
        if (offsets.Length <= 1)
        {
            return 0;
        }

        int index = Array.BinarySearch(offsets, Math.Max(0, value));
        if (index >= 0)
        {
            return Math.Clamp(index == 0 ? 0 : index - 1, 0, offsets.Length - 2);
        }

        return Math.Clamp((~index) - 1, 0, offsets.Length - 2);
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
            // At/near the bottom (or content shorter than the viewport) → add enough rows to span the
            // empty gap below the content (large after a zoom-out), plus a buffer to scroll into.
            double gap = (offsetY + viewport) - extent;
            int fill = gap > 0 ? (int)Math.Ceiling(gap / Math.Max(1, DefaultRowHeight * this.zoom)) : 0;
            target = this.sheet.RowCount + fill + RowBuffer;
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
            // At/near the right (or content narrower than the viewport) → add enough columns to span
            // the empty gap to the right (large after a zoom-out), plus a buffer to scroll into.
            double gap = (offsetX + viewport) - extent;
            int fill = gap > 0 ? (int)Math.Ceiling(gap / Math.Max(1, this.sheet.DefaultColumnWidth * this.zoom)) : 0;
            target = this.sheet.ColumnCount + fill + ColumnBuffer;
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
            // Positions/offsets are in the (zoom-scaled) layout; column/row offsets already include zoom
            // via OffsetToColumn/Row, and the logical pixel offsets/sizes are scaled by zoom so charts
            // resize together with the cells.
            double left = this.OffsetToColumn(chart.AnchorColumn) + (chart.OffsetX * this.zoom);
            double top = this.OffsetToRow(chart.AnchorRow) + (chart.OffsetY * this.zoom);

            // The displayed size spans from→to cells in OUR (scaled) layout; fall back to the stored
            // pixel size for one-cell-anchored charts.
            double z = this.zoom > 0 ? this.zoom : 1;
            double scaledWidth;
            double scaledHeight;
            if (chart.ToColumn is { } toCol && chart.ToRow is { } toRow)
            {
                scaledWidth = this.OffsetToColumn(toCol) + (chart.ToOffsetX * this.zoom) - left;
                scaledHeight = this.OffsetToRow(toRow) + (chart.ToOffsetY * this.zoom) - top;
            }
            else
            {
                scaledWidth = chart.Width * this.zoom;
                scaledHeight = chart.Height * this.zoom;
            }

            // Build the chart at its LOGICAL (zoom-independent) size and scale the whole card with a
            // ScaleTransform — so zoom shrinks/grows the chart as a fixed object (like Excel) instead of
            // re-laying-out its contents (labels, axis) at every zoom level.
            double width = Math.Clamp(scaledWidth / z, 120, 2400);
            double height = Math.Clamp(scaledHeight / z, 80, 1600);

            (List<string> cats, List<ChartSeriesData> ser) = this.ExtractChartData(chart);
            Control? control = ChartFactory.Build(chart.Kind, cats, ser, width, height, chart.ShowDataLabels,
                chart.Title, chart.ShowLegend, chart.ValueAxisFormat, chart.AxisMinimum, chart.AxisMaximum, chart.AxisMajorStep, chart.LegendPosition);
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
            ToolTip.SetTip(settings, Get(S.EditChart));
            SheetChart captured = chart;
            settings.Click += (_, _) => this.ChartEditRequested?.Invoke(captured);

            Grid content = new();
            content.Children.Add(control);
            content.Children.Add(settings);

            bool selected = ReferenceEquals(chart, this.selectedChart);
            Border card = new()
            {
                Background = Brushes.White,
                BorderBrush = selected ? SelectionBorder : SheetGridLine,
                BorderThickness = new Thickness(selected ? 2 : 1),
                Child = content,
                Width = width,
                Height = height,
                Cursor = new Cursor(StandardCursorType.SizeAll),
                // Scale the whole (logical-size) card by the zoom factor, anchored at its top-left so it
                // stays put — the chart scales uniformly as a fixed object rather than re-laying-out.
                RenderTransform = new ScaleTransform(z, z),
                RenderTransformOrigin = RelativePoint.TopLeft,
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
                    // cardWidth/cardHeight are already logical (the card is scaled via a ScaleTransform).
                    this.sheet?.MoveChart(captured, row, col, cardWidth, cardHeight);
                }
                pressed = false;
                moved = false;
            };

            // Right-click menu with a Delete item (acts on this chart directly).
            MenuItem deleteItem = new() { Header = Get(S.DeleteChart) };
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
            ToolTip.SetTip(grip, Get(S.DragToResize));
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
                // The pointer delta is in scaled space; the card is logical-size (scaled by transform),
                // so convert the delta back to logical before applying it.
                double zoomNow = this.zoom > 0 ? this.zoom : 1;
                double w = Math.Clamp(startW + ((p.X - startSizePointer.X) / zoomNow), 120, 2400);
                double h = Math.Clamp(startH + ((p.Y - startSizePointer.Y) / zoomNow), 80, 1600);
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
                    // card.Width/Height are already logical (the card is scaled via a ScaleTransform).
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
            // A selected chart's card carries a thick blue selection border — reset it to the normal
            // gridline look so the snapshot (and the print/PDF) never shows the selection highlight.
            IBrush? borderBrush = card.BorderBrush;
            Thickness borderThickness = card.BorderThickness;
            card.BorderBrush = SheetGridLine;
            card.BorderThickness = new Thickness(1);
            ITransform? transform = card.RenderTransform;
            card.RenderTransform = null; // snapshot at the card's logical size, not the zoomed size
            try
            {
                PixelSize pixelSize = new((int)Math.Ceiling(width), (int)Math.Ceiling(height));
                using RenderTargetBitmap bitmap = new(pixelSize, new Vector(96, 96));
                bitmap.Render(card);
                using MemoryStream stream = new();
                bitmap.Save(stream, new PngBitmapEncoderOptions());
                // card.Width/Height are logical; only the position needs converting from scaled to
                // logical grid pixels (the PDF exporter lays cells out at the sheet's logical sizes).
                double z = this.zoom > 0 ? this.zoom : 1;
                images.Add(new ChartImage(stream.ToArray(), left / z, top / z, width, height));
            }
            catch
            {
                // Skip a chart that can't be snapshotted rather than failing the whole export.
            }
            finally
            {
                card.RenderTransform = transform;
                card.BorderBrush = borderBrush;
                card.BorderThickness = borderThickness;
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
        if (this.sheet is null)
        {
            return Array.Empty<ChartWriteData>();
        }

        return BuildChartWriteData(this.sheet, sheetName);
    }

    internal static IReadOnlyList<ChartWriteData> BuildChartWriteData(
        SpreadsheetViewModel sheet,
        string sheetName)
    {
        List<ChartWriteData> result = new();
        foreach (SheetChart chart in sheet.Charts)
        {
            (List<string> categories, List<ChartSeriesData> series) = ExtractChartData(sheet, chart);

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
            double left = LogicalColumnOffset(sheet, chart.AnchorColumn) + chart.OffsetX;
            double top = LogicalRowOffset(sheet, chart.AnchorRow) + chart.OffsetY;
            double width = chart.Width;
            double height = chart.Height;
            if (chart.ToColumn is { } toColumn && chart.ToRow is { } toRow)
            {
                width = LogicalColumnOffset(sheet, toColumn) + chart.ToOffsetX - left;
                height = LogicalRowOffset(sheet, toRow) + chart.ToOffsetY - top;
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

    private static double LogicalColumnOffset(SpreadsheetViewModel sheet, int column) =>
        sheet.Columns.Take(Math.Clamp(column, 0, sheet.ColumnCount)).Sum(static item => item.Width);

    private static double LogicalRowOffset(SpreadsheetViewModel sheet, int row) =>
        sheet.Rows.Take(Math.Clamp(row, 0, sheet.RowCount)).Sum(static item =>
            item.Height ?? Math.Max(DefaultRowHeight, item.Cells.DefaultIfEmpty().Max(static cell => cell?.FontSize ?? CellViewModel.DefaultFontSize) + 8));

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
    private (List<string> Categories, List<ChartSeriesData> Series) ExtractChartData(SheetChart chart) =>
        ExtractChartData(this.sheet!, chart);

    private static (List<string> Categories, List<ChartSeriesData> Series) ExtractChartData(
        SpreadsheetViewModel sheet,
        SheetChart chart)
    {
        string Display(CellAddress a) => sheet.GetCell(a)?.DisplayText ?? string.Empty;

        // True when the range resolves to at least one non-empty cell in the loaded sheet, i.e. the
        // chart's data lives here and should be read live (so edits/recalcs update the chart). When
        // false (data on another sheet or outside the loaded range) we fall back to the file snapshot.
        bool RangeHasData(CellRange range)
        {
            if (range.RowSpan <= 0 || range.ColumnSpan <= 0)
            {
                return false;
            }

            return EnumerateRange(sheet, range).Any(a => !string.IsNullOrWhiteSpace(Display(a)));
        }

        List<ChartSeriesData> series = new();
        foreach (ChartSeriesSpec spec in chart.Series)
        {
            // A series can span several ranges/cells — concatenate their cells in order into one series.
            List<double> values = spec.Ranges.Any(RangeHasData)
                ? spec.Ranges.SelectMany(range => EnumerateRange(sheet, range)).Select(a => ParseNumber(Display(a)) ?? double.NaN).ToList()
                : spec.CachedValues?.ToList() ?? new List<double>();
            series.Add(new ChartSeriesData(spec.Title, values));
        }

        int maxLen = series.Count == 0 ? 0 : series.Max(s => s.Values.Count);
        List<string> categories;
        if (chart.Categories is { } cats && RangeHasData(cats))
        {
            categories = EnumerateRange(sheet, cats).Select(Display).ToList();
        }
        else if (chart.CachedCategories is { Count: > 0 } cached)
        {
            categories = cached.ToList();
        }
        else if (chart.Categories is { } cats2)
        {
            categories = EnumerateRange(sheet, cats2).Select(Display).ToList();
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
    private static IEnumerable<CellAddress> EnumerateRange(SpreadsheetViewModel sheet, CellRange range)
    {
        int rowEnd = Math.Min(range.Row + range.RowSpan - 1, sheet.RowCount - 1);
        int colEnd = Math.Min(range.Column + range.ColumnSpan - 1, sheet.ColumnCount - 1);
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

    private Border HeaderCell(string text)
    {
        return new Border
        {
            Background = this.headerBackground,
            BorderBrush = this.headerLine,
            BorderThickness = new Thickness(0, 0, 1, 1),
            Child = new TextBlock
            {
                Text = text,
                Foreground = this.headerText,
                FontSize = 12 * this.zoom,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
            },
        };
    }

    private void PlaceColumnHeader(Control control, int column)
    {
        Canvas.SetLeft(control, this.OffsetToColumn(column));
        Canvas.SetTop(control, 0);
        control.Width = this.columnWidths[column];
        control.Height = ColumnHeaderHeight * this.zoom;
    }

    private void PlaceRowHeader(Control control, int row)
    {
        Canvas.SetLeft(control, 0);
        Canvas.SetTop(control, this.OffsetToRow(row));
        control.Width = RowHeaderWidth * this.zoom;
        control.Height = this.rowHeights[row];
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
            Height = ColumnHeaderHeight * this.zoom,
            Cursor = new Cursor(StandardCursorType.SizeWestEast),
            ContextFlyout = this.BuildColumnMenu(column),
        };
        Canvas.SetLeft(handle, this.OffsetToColumn(column + 1) - 3);
        Canvas.SetTop(handle, 0);
        handle.PointerPressed += (_, e) =>
        {
            if (!e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed)
            {
                if (e.GetCurrentPoint(handle).Properties.IsRightButtonPressed)
                {
                    this.SelectColumn(column);
                }
                return;
            }
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
            Width = RowHeaderWidth * this.zoom,
            Height = 6,
            Cursor = new Cursor(StandardCursorType.SizeNorthSouth),
            ContextFlyout = this.BuildRowMenu(row),
        };
        Canvas.SetLeft(handle, 0);
        Canvas.SetTop(handle, this.OffsetToRow(row + 1) - 3);
        handle.PointerPressed += (_, e) =>
        {
            if (!e.GetCurrentPoint(handle).Properties.IsLeftButtonPressed)
            {
                if (e.GetCurrentPoint(handle).Properties.IsRightButtonPressed)
                {
                    this.SelectRow(row);
                }
                return;
            }
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
        for (int c = column; c < this.columnWidths.Length; c++)
        {
            this.columnOffsets[c + 1] = this.columnOffsets[c] + this.columnWidths[c];
        }
        this.UpdateCanvasExtents();
        this.renderedViewport = null;
        this.ScheduleViewportRender();
        if (this.sheet is not null && column < this.sheet.ColumnCount)
        {
            // Store the logical (unscaled) width so zoom doesn't leak into the model / saved file.
            this.sheet.Columns[column].Width = this.zoom > 0 ? width / this.zoom : width;
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
        for (int r = row; r < this.rowHeights.Length; r++)
        {
            this.rowOffsets[r + 1] = this.rowOffsets[r] + this.rowHeights[r];
        }
        this.UpdateCanvasExtents();
        this.renderedViewport = null;
        this.ScheduleViewportRender();
        if (this.sheet is not null && row < this.sheet.RowCount)
        {
            // Store the logical (unscaled) height so zoom doesn't leak into the model.
            this.sheet.Rows[row].Height = this.zoom > 0 ? height / this.zoom : height;
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
        int rightRun = align != TextAlignment.Right ? this.EmptyRun(row, column, +1) : 0;
        int leftRun = align != TextAlignment.Left ? this.EmptyRun(row, column, -1) : 0;
        if (rightRun == 0 && leftRun == 0)
        {
            return (0, 0);
        }

        // Only overflow as far as the text is actually wide — the gridlines that vanish are just the
        // ones under the overflowing text, not every empty cell to the side. If the text fits its own
        // cell there is no spill (and every gridline stays).
        double overflow = this.MeasureTextWidth(cell) - (this.columnWidths[column] - CellTextPadding);
        if (overflow <= 0)
        {
            return (0, 0);
        }

        return align switch
        {
            TextAlignment.Right => (this.ColumnsToCover(column, -1, overflow, leftRun), 0),
            TextAlignment.Center => (this.ColumnsToCover(column, -1, overflow / 2, leftRun), this.ColumnsToCover(column, +1, overflow / 2, rightRun)),
            _ => (0, this.ColumnsToCover(column, +1, overflow, rightRun)),
        };
    }

    // Total horizontal padding (both sides) of a cell's text, matching the in-cell/overlay Margin.
    private const double CellTextPadding = 8;

    // Measures the rendered width (px, at the current zoom) of a cell's text, summing per-run widths
    // so bold/italic runs are measured with the right typeface.
    private double MeasureTextWidth(CellViewModel cell)
    {
        FontFamily family = string.IsNullOrWhiteSpace(cell.FontFamily) ? FontFamily.Default : new FontFamily(cell.FontFamily);
        double size = cell.FontSize * this.zoom;
        double width = 0;
        foreach (CellRun run in cell.EffectiveRuns)
        {
            if (string.IsNullOrEmpty(run.Text))
            {
                continue;
            }
            Typeface typeface = new(family, run.Italic ? FontStyle.Italic : FontStyle.Normal, run.Bold ? FontWeight.Bold : FontWeight.Normal);
            FormattedText formatted = new(run.Text, CultureInfo.CurrentCulture, FlowDirection.LeftToRight, typeface, size, null);
            width += formatted.Width;
        }

        return width;
    }

    // How many consecutive columns (from `column`, in `direction`) the overflowing text reaches into:
    // it crosses a column's boundary — and so removes that gridline — until the accumulated width of
    // the columns entered covers the overflow. Capped by the empty run so text never spills onto data.
    private int ColumnsToCover(int column, int direction, double overflow, int maxRun)
    {
        int covered = 0;
        double accumulated = 0;
        for (int i = 1; i <= maxRun; i++)
        {
            int col = column + (direction * i);
            if (col < 0 || col >= this.columnWidths.Length)
            {
                break;
            }
            covered = i;
            accumulated += this.columnWidths[col];
            if (accumulated >= overflow)
            {
                break;
            }
        }

        return covered;
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
            FontSize = cell.FontSize * this.zoom,
            TextAlignment = cell.TextAlignment,
            // The sheet stays light in dark mode, so cell text must be explicitly dark — the
            // inherited theme foreground would be near-white. Per-run colors still override this.
            Foreground = CellText,
            IsHitTestVisible = false,
        };
        if (!string.IsNullOrWhiteSpace(cell.FontFamily))
        {
            text.FontFamily = new FontFamily(cell.FontFamily);
        }
        RenderInlines(text, cell, this.FindCellValueRule(row, startColumn, cell));
        this.PlaceSheetControl(text, row, startColumn, 1, span);
        return text;
    }

    private Control BuildCell(CellViewModel cell, int row, int column, int rowSpan, int columnSpan, bool spilling = false, bool omitRightBorder = false)
    {
        TextBlock text = new()
        {
            VerticalAlignment = VerticalAlignment.Center,
            Margin = new Thickness(4, 2),
            // When the text spills into empty neighbours, an overlay TextBlock draws the full text on
            // top; the in-cell one then shouldn't show an ellipsis (the overlay covers it).
            TextTrimming = spilling ? TextTrimming.None : TextTrimming.CharacterEllipsis,
            FontSize = cell.FontSize * this.zoom,
            TextAlignment = cell.TextAlignment,
            // The sheet stays light in dark mode, so cell text must be explicitly dark — the
            // inherited theme foreground would be near-white. Per-run colors still override this.
            Foreground = CellText,
        };
        if (!string.IsNullOrWhiteSpace(cell.FontFamily))
        {
            text.FontFamily = new FontFamily(cell.FontFamily);
        }
        // A spilling cell's text is drawn entirely by the overlay; keep the in-cell one empty so the
        // two don't render on top of each other.
        if (!spilling)
        {
            RenderInlines(text, cell, this.FindCellValueRule(row, column, cell));
        }

        // Conditional-formatting data bar: a proportional bar behind (or instead of) the value.
        SheetDataBar? dataBar = this.FindDataBar(row, column);
        Grid? barGrid = null;
        if (dataBar is { } db)
        {
            barGrid = new Grid();
            Border fill = new() { Background = this.DataBarBrush(db), Margin = new Thickness(1) };
            Grid.SetColumn(fill, 0);
            barGrid.Children.Add(fill);
            SetBarFraction(barGrid, db, cell);
            if (!db.ShowValue)
            {
                text.Inlines?.Clear(); // show only the bar (Excel "show bar only")
            }
        }

        Control cellContent = barGrid is null ? text : Layered(barGrid, text);
        Border? filterButton = null;
        if (this.sheet?.AutoFilterRange is { } filterRange &&
            row == filterRange.Row && column >= filterRange.Column &&
            column < filterRange.Column + filterRange.ColumnSpan)
        {
            text.Margin = new Thickness(4, 2, 23 * this.zoom, 2);
            filterButton = new Border
            {
                Width = 17 * this.zoom,
                Height = Math.Min(19 * this.zoom, this.rowHeights[row] - 2),
                Margin = new Thickness(0, 0, 2 * this.zoom, 0),
                HorizontalAlignment = HorizontalAlignment.Right,
                VerticalAlignment = VerticalAlignment.Center,
                Background = Brushes.White,
                BorderBrush = SheetGridLine,
                BorderThickness = new Thickness(1),
                Focusable = true,
                Cursor = new Cursor(StandardCursorType.Hand),
                Child = new TextBlock
                {
                    Text = "▼",
                    Foreground = CellText,
                    FontSize = 9 * this.zoom,
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Center,
                    IsHitTestVisible = false,
                },
            };
            Border button = filterButton;
            button.PointerPressed += (_, e) =>
            {
                if (e.GetCurrentPoint(button).Properties.IsLeftButtonPressed)
                {
                    this.BuildColumnMenu(column).ShowAt(button);
                    e.Handled = true;
                }
            };
            button.KeyDown += (_, e) =>
            {
                if (e.Key is Key.Enter or Key.Space or Key.Down)
                {
                    this.BuildColumnMenu(column).ShowAt(button);
                    e.Handled = true;
                }
            };
            ToolTip.SetTip(button, Format(S.SortOrFilter, cell.DisplayText));
            cellContent = Layered(cellContent, button);
        }
        Border border = new()
        {
            BorderBrush = SheetGridLine,
            // A spilling cell drops the gridline between it and the empty neighbour it overflows into,
            // so the overlaid text reads as continuous. The cell keeps its own hit-testable box.
            BorderThickness = new Thickness(0, 0, omitRightBorder ? 0 : 1, 1),
            Background = cell.Background ?? Brushes.Transparent,
            Child = cellContent,
        };
        if (filterButton is not null)
        {
            border.ContextFlyout = this.BuildColumnMenu(column);
        }
        Grid? annotatedContent = null;
        bool hasAnnotationTip = false;
        void UpdateAnnotation()
        {
            if (!string.IsNullOrWhiteSpace(cell.Comment) && annotatedContent is null)
            {
                border.Child = null;
                annotatedContent = new Grid();
                annotatedContent.Children.Add(cellContent);
                annotatedContent.Children.Add(new Polygon
                {
                    Points = new Points { new Point(0, 0), new Point(9, 0), new Point(9, 9) },
                    Fill = Brushes.Gold,
                    HorizontalAlignment = HorizontalAlignment.Right,
                    VerticalAlignment = VerticalAlignment.Top,
                    IsHitTestVisible = false,
                });
                border.Child = annotatedContent;
            }
            else if (string.IsNullOrWhiteSpace(cell.Comment) && annotatedContent is not null)
            {
                border.Child = null;
                annotatedContent.Children.Remove(cellContent);
                annotatedContent = null;
                border.Child = cellContent;
            }

            string? tip = string.IsNullOrWhiteSpace(cell.CommentAuthor) || string.IsNullOrWhiteSpace(cell.Comment)
                ? cell.Comment
                : $"{cell.CommentAuthor}: {cell.Comment}";
            if (!string.IsNullOrWhiteSpace(cell.Hyperlink))
            {
                tip = string.IsNullOrWhiteSpace(tip)
                    ? Format(S.CtrlClickOpen, cell.Hyperlink)
                    : $"{tip}\n\n{Format(S.CtrlClickOpen, cell.Hyperlink)}";
            }
            if (tip is not null || hasAnnotationTip)
            {
                ToolTip.SetTip(border, tip);
                hasAnnotationTip = tip is not null;
            }
        }
        UpdateAnnotation();
        this.PlaceSheetControl(border, row, column, rowSpan, columnSpan);

        CellAddress address = new(row, column);
        border.PointerPressed += (_, e) =>
        {
            if (filterButton is not null && IsWithin(e.Source as Visual, filterButton))
            {
                return;
            }
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
            if (e.KeyModifiers.HasFlag(KeyModifiers.Control) && !string.IsNullOrWhiteSpace(cell.Hyperlink))
            {
                this.OpenHyperlink(cell.Hyperlink);
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
                    if (this.applyingCellFormat)
                    {
                        break;
                    }
                    // The editor already paints the live text. Defer the comparatively expensive
                    // spill calculation/rebuild until the editing session is committed.
                    if (this.editingSession)
                    {
                        break;
                    }

                    // The text changed, so the spill extent may have too. Rebuild when the cell spills
                    // now OR spilled before — a spilling cell's text lives in a separate overlay whose
                    // span (and the neighbour gridlines it clears) is only recomputed on a rebuild.
                    // This fires exactly when the new value lands on the model, so a cell that just grew
                    // past its width starts overflowing even after being cleared and retyped.
                    if (spilling || (this.spans is not null && this.spans[row, column] is (1, 1) && this.ComputeSpill(row, column, cell) is not (0, 0)))
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
                        RenderInlines(text, cell, this.FindCellValueRule(row, column, cell));
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
                case nameof(CellViewModel.Comment) or nameof(CellViewModel.CommentAuthor) or nameof(CellViewModel.Hyperlink):
                    UpdateAnnotation();
                    if (e.PropertyName == nameof(CellViewModel.Hyperlink) && spilling)
                    {
                        this.Rebuild(); // the visible text lives in a separate spill overlay
                    }
                    else if (e.PropertyName == nameof(CellViewModel.Hyperlink))
                    {
                        RenderInlines(text, cell, this.FindCellValueRule(row, column, cell));
                    }
                    break;
                case nameof(CellViewModel.FontSize):
                    if (!this.applyingCellFormat)
                    {
                        this.Rebuild();
                    }
                    break;
            }
        };
        cell.PropertyChanged += handler;
        this.cellSubscriptions.Add((cell, handler));
        return border;
    }

    private void PlaceSheetControl(Control control, int row, int column, int rowSpan, int columnSpan)
    {
        int lastRow = Math.Clamp(row + Math.Max(1, rowSpan), 0, this.rowOffsets.Length - 1);
        int lastColumn = Math.Clamp(column + Math.Max(1, columnSpan), 0, this.columnOffsets.Length - 1);
        double left = this.OffsetToColumn(column);
        double top = this.OffsetToRow(row);
        Canvas.SetLeft(control, left);
        Canvas.SetTop(control, top);
        control.Width = Math.Max(0, this.OffsetToColumn(lastColumn) - left);
        control.Height = Math.Max(0, this.OffsetToRow(lastRow) - top);
    }

    private SheetCellValueRule? FindCellValueRule(int row, int column, CellViewModel cell)
    {
        double? value = cell.NumericValue ?? ParseNumber(cell.DisplayText);
        if (value is null || this.sheet is null)
        {
            return null;
        }
        foreach (SheetCellValueRule rule in this.sheet.CellValueRules)
        {
            CellRange range = rule.Range;
            if (row >= range.Row && row < range.Row + range.RowSpan &&
                column >= range.Column && column < range.Column + range.ColumnSpan &&
                rule.Matches(value.Value))
            {
                return rule;
            }
        }
        return null;
    }

    private IBrush DataBarBrush(SheetDataBar bar)
    {
        if (!bar.Gradient)
        {
            return new SolidColorBrush(bar.Color);
        }
        if (this.gradientBarBrushes.TryGetValue(bar.Color, out IBrush? cached))
        {
            return cached;
        }
        Color pale = Color.FromRgb(
            (byte)(bar.Color.R + ((255 - bar.Color.R) * 0.82)),
            (byte)(bar.Color.G + ((255 - bar.Color.G) * 0.82)),
            (byte)(bar.Color.B + ((255 - bar.Color.B) * 0.82)));
        LinearGradientBrush brush = new()
        {
            StartPoint = new RelativePoint(0, 0, RelativeUnit.Relative),
            EndPoint = new RelativePoint(1, 0, RelativeUnit.Relative),
            GradientStops = new GradientStops
            {
                new GradientStop(bar.Color, 0),
                new GradientStop(pale, 1),
            },
        };
        this.gradientBarBrushes[bar.Color] = brush;
        return brush;
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
        double value = cell.NumericValue ?? ParseNumber(cell.DisplayText) ?? bar.Min;
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

    private static void RenderInlines(TextBlock target, CellViewModel cell, SheetCellValueRule? valueRule = null)
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
                FontWeight = (valueRule?.Bold ?? r.Bold) ? FontWeight.Bold : FontWeight.Normal,
                FontStyle = r.Italic ? FontStyle.Italic : FontStyle.Normal,
            };
            if (r.Underline || r.Strikethrough || !string.IsNullOrWhiteSpace(cell.Hyperlink))
            {
                TextDecorationCollection decorations = new();
                if (r.Underline || !string.IsNullOrWhiteSpace(cell.Hyperlink))
                {
                    decorations.Add(new TextDecoration { Location = TextDecorationLocation.Underline });
                }

                if (r.Strikethrough)
                {
                    decorations.Add(new TextDecoration { Location = TextDecorationLocation.Strikethrough });
                }

                run.TextDecorations = decorations;
            }

            if (valueRule?.Foreground is { } conditionalColor)
            {
                run.Foreground = new SolidColorBrush(conditionalColor);
            }
            else if (r.Foreground is { } fg)
            {
                run.Foreground = new SolidColorBrush(fg);
            }
            else if (!string.IsNullOrWhiteSpace(cell.Hyperlink))
            {
                run.Foreground = LinkText;
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

    private void SetFillPart(Border part, int row, int col, int rowSpan, int colSpan, bool show)
    {
        if (show && rowSpan > 0 && colSpan > 0)
        {
            this.PlaceSheetControl(part, row, col, rowSpan, colSpan);
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
            this.SetFillPart(this.fillParts[0], selRow, selCol, ar1 - selRow, selColSpan, ar1 > selRow);             // above anchor
            this.SetFillPart(this.fillParts[1], ar2 + 1, selCol, r2 - ar2, selColSpan, ar2 < r2);                    // below anchor
            this.SetFillPart(this.fillParts[2], ar1, selCol, ar2 - ar1 + 1, ac1 - selCol, ac1 > selCol);             // left of anchor
            this.SetFillPart(this.fillParts[3], ar1, ac2 + 1, ar2 - ar1 + 1, c2 - ac2, ac2 < c2);                   // right of anchor

            this.PlaceSheetControl(this.anchorBorder, ar1, ac1, rowSpan, colSpan);
            this.anchorBorder.IsVisible = true;
        }
        this.PlaceSheetControl(this.selection, selRow, selCol, selRowSpan, selColSpan);
        this.selection.IsVisible = true;

        // Position the shared editor over the anchor cell and load that cell's content.
        this.PlaceSheetControl(this.editorHost, row, col, rowSpan, colSpan);
        CellViewModel? cell = this.sheet.GetCell(new CellAddress(row, col));
        // Opaque background (cell fill, else white) so the editor covers the display text beneath it.
        this.editorHost.Background = cell?.Background ?? SheetSurface;
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

    private void SelectRow(int row)
    {
        if (this.sheet is null)
        {
            return;
        }
        this.editorSuspended = true;
        this.CellActivated?.Invoke(new CellAddress(row, 0));
        this.rangeAnchor = new CellAddress(row, 0);
        this.rangeFocus = new CellAddress(row, this.sheet.ColumnCount - 1);
        this.UpdateSelection();
    }

    private void SelectColumn(int column)
    {
        if (this.sheet is null)
        {
            return;
        }
        this.editorSuspended = true;
        this.CellActivated?.Invoke(new CellAddress(0, column));
        this.rangeAnchor = new CellAddress(0, column);
        this.rangeFocus = new CellAddress(this.sheet.RowCount - 1, column);
        this.UpdateSelection();
    }

    private void InsertSelectedRows(int row, int count = 1)
    {
        if (this.ReadOnly || this.sheet is null)
        {
            return;
        }
        this.sheet.InsertRows(row, count);
        this.Rebuild();
        this.SelectCell(new CellAddress(row, 0));
    }

    private void DeleteSelectedRows(int row, int count = 1)
    {
        if (this.ReadOnly || this.sheet is null)
        {
            return;
        }
        this.sheet.DeleteRows(row, count);
        this.Rebuild();
        this.SelectCell(new CellAddress(Math.Min(row, this.sheet.RowCount - 1), 0));
    }

    private MenuFlyout BuildRowMenu(int row)
    {
        MenuFlyout menu = new();
        MenuItem insert = new() { Header = Get(S.InsertRowAbove) };
        insert.Click += (_, _) => this.InsertSelectedRows(row);
        MenuItem delete = new() { Header = Get(S.DeleteRow) };
        delete.Click += (_, _) => this.DeleteSelectedRows(row);
        menu.Items.Add(insert);
        menu.Items.Add(delete);
        return menu;
    }

    private MenuFlyout BuildColumnMenu(int column)
    {
        MenuFlyout menu = new();
        if (this.sheet is null)
        {
            return menu;
        }
        int headerRow = this.sheet.LikelyHeaderRow();
        MenuItem ascending = new() { Header = Get(S.SortAscending) };
        ascending.Click += (_, _) => this.SortColumn(column, headerRow, descending: false);
        MenuItem descending = new() { Header = Get(S.SortDescending) };
        descending.Click += (_, _) => this.SortColumn(column, headerRow, descending: true);
        menu.Items.Add(ascending);
        menu.Items.Add(descending);
        menu.Items.Add(new Separator());

        MenuItem filter = new() { Header = Get(S.FilterByValue) };
        filter.Items.Add(new MenuItem { Header = Get(S.ChooseValue) });
        filter.SubmenuOpened += (_, _) =>
        {
            filter.Items.Clear();
            if (this.sheet is null)
            {
                return;
            }
            foreach (string value in this.sheet.ColumnValues(column, headerRow))
            {
                string selectedValue = value;
                MenuItem option = new() { Header = value.Length == 0 ? Get(S.Blanks) : value };
                option.Click += (_, _) => this.FilterColumn(column, selectedValue, headerRow);
                filter.Items.Add(option);
            }
            filter.Items.Add(new Separator());
            MenuItem custom = new() { Header = Get(S.EnterValue) };
            custom.Click += async (_, _) =>
            {
                string? value = await this.PromptCellTextAsync(
                    Get(S.FilterByValue), Get(S.FilterWhereEquals), string.Empty);
                if (value is not null)
                {
                    this.FilterColumn(column, value, headerRow);
                }
            };
            filter.Items.Add(custom);
        };
        menu.Items.Add(filter);
        if (this.sheet.HasColumnFilter(column))
        {
            MenuItem clear = new() { Header = Get(S.ClearColumnFilter) };
            clear.Click += (_, _) => this.FilterColumn(column, null, headerRow);
            menu.Items.Add(clear);
        }
        MenuItem showAll = new() { Header = Get(S.ShowAllRows) };
        showAll.Click += (_, _) =>
        {
            if (this.ReadOnly || this.sheet is null)
            {
                return;
            }
            this.sheet.ShowAllRows();
            this.Rebuild();
        };
        menu.Items.Add(showAll);
        return menu;
    }

    private void SortColumn(int column, int headerRow, bool descending)
    {
        if (this.ReadOnly || this.sheet is null)
        {
            return;
        }
        this.sheet.SortRows(column, headerRow, descending);
        this.Rebuild();
        this.SelectCell(this.sheet.ActiveCell);
    }

    private void FilterColumn(int column, string? value, int headerRow)
    {
        if (this.ReadOnly || this.sheet is null)
        {
            return;
        }
        this.sheet.FilterColumn(column, value, headerRow);
        this.Rebuild();
    }

    private async Task<string?> PromptCellTextAsync(string title, string caption, string current,
        bool multiline = false, Func<string, bool>? valid = null, string? validationMessage = null)
    {
        if (TopLevel.GetTopLevel(this) is not Window owner)
        {
            return null;
        }

        TextBox input = new()
        {
            Text = current,
            AcceptsReturn = multiline,
            TextWrapping = multiline ? TextWrapping.Wrap : TextWrapping.NoWrap,
            MinHeight = multiline ? 100 : 0,
        };
        TextBlock error = new() { Foreground = Brushes.Firebrick, IsVisible = false, Text = validationMessage };
        Button save = new() { Content = Get(S.Save), IsDefault = true, MinWidth = 80 };
        Button cancel = new() { Content = Get(S.Cancel), IsCancel = true, MinWidth = 80 };
        StackPanel buttons = new() { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8 };
        buttons.Children.Add(cancel);
        buttons.Children.Add(save);
        StackPanel panel = new() { Margin = new Thickness(16), Spacing = 10 };
        panel.Children.Add(new TextBlock { Text = caption, TextWrapping = TextWrapping.Wrap });
        panel.Children.Add(input);
        panel.Children.Add(error);
        panel.Children.Add(buttons);
        Window dialog = new()
        {
            Title = title,
            Width = 450,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Content = panel,
        };
        save.Click += (_, _) =>
        {
            string value = input.Text?.Trim() ?? string.Empty;
            if (valid is not null && value.Length > 0 && !valid(value))
            {
                error.IsVisible = true;
                return;
            }
            dialog.Close(value);
        };
        cancel.Click += (_, _) => dialog.Close(null);
        dialog.Opened += (_, _) => { input.Focus(); input.SelectAll(); };
        return await dialog.ShowDialog<string?>(owner);
    }

    private async Task EditCommentAsync()
    {
        if (this.ReadOnly || this.sheet is not { } sheet || this.ActiveCell() is not { } cell)
        {
            return;
        }
        this.editorSuspended = true;
        string? comment = await this.PromptCellTextAsync(
            Get(string.IsNullOrWhiteSpace(cell.Comment) ? S.InsertComment : S.EditComment),
            Get(S.CommentPrompt), cell.Comment ?? string.Empty, multiline: true);
        if (comment is null || comment == cell.Comment)
        {
            return;
        }
        sheet.Checkpoint();
        cell.Comment = comment.Length == 0 ? null : comment;
        cell.CommentAuthor = comment.Length == 0 ? null : cell.CommentAuthor ?? Get(S.CommentAuthor);
        sheet.NotifyMetadataChanged();
    }

    private static bool IsValidLinkTarget(string target) =>
        target.StartsWith('#') && target.Length > 1 ||
        Uri.TryCreate(target, UriKind.Absolute, out Uri? uri) &&
        uri.Scheme is "http" or "https" or "mailto";

    private async Task EditHyperlinkAsync()
    {
        if (this.ReadOnly || this.sheet is not { } sheet || this.ActiveCell() is not { } cell)
        {
            return;
        }
        this.editorSuspended = true;
        string? target = await this.PromptCellTextAsync(
            Get(S.Hyperlink), Get(S.HyperlinkPrompt),
            cell.Hyperlink ?? string.Empty, valid: IsValidLinkTarget,
            validationMessage: Get(S.HyperlinkInvalid));
        if (target is null || target == cell.Hyperlink)
        {
            return;
        }
        sheet.Checkpoint();
        if (target.Length > 0 && string.IsNullOrEmpty(cell.Text))
        {
            cell.Text = target;
        }
        cell.Hyperlink = target.Length == 0 ? null : target;
        sheet.NotifyMetadataChanged();
    }

    private async Task FormatCellsAsync()
    {
        if (this.ReadOnly || this.sheet is not { } sheet || this.ActiveCell() is not { } active ||
            TopLevel.GetTopLevel(this) is not Window owner)
        {
            return;
        }
        this.editorSuspended = true;

        static TextBox Field(string? value) => new() { Text = value ?? string.Empty };
        static string Hex(Color? color) => color is { } c ? $"#{c.R:X2}{c.G:X2}{c.B:X2}" : string.Empty;
        StackPanel fields = new() { Spacing = 7 };
        void AddField(string label, Control control)
        {
            fields.Children.Add(new TextBlock { Text = label });
            fields.Children.Add(control);
        }

        TextBox numberFormat = Field(active.NumberFormat);
        AddField(Get(S.NumberFormatLabel), numberFormat);
        ComboBox alignment = new()
        {
            ItemsSource = new[] { Get(S.Left), Get(S.Center), Get(S.Right) },
            SelectedIndex = active.TextAlignment switch
            {
                TextAlignment.Center => 1,
                TextAlignment.Right => 2,
                _ => 0,
            }
        };
        AddField(Get(S.Alignment), alignment);
        TextBox fontName = Field(active.FontFamily);
        AddField(Get(S.FontFamilyLabel), fontName);
        TextBox fontSize = Field((active.FontSize * 3.0 / 4.0).ToString("0.##", CultureInfo.InvariantCulture));
        AddField(Get(S.FontSizePoints), fontSize);
        CheckBox bold = new() { Content = Get(S.Bold), IsChecked = active.IsAllBold };
        CheckBox italic = new() { Content = Get(S.Italic), IsChecked = active.IsAllItalic };
        CheckBox underline = new() { Content = Get(S.Underline), IsChecked = active.IsAllUnderline };
        CheckBox strike = new() { Content = Get(S.Strikethrough), IsChecked = active.IsAllStrikethrough };
        StackPanel flags = new() { Orientation = Orientation.Horizontal, Spacing = 12 };
        flags.Children.Add(bold);
        flags.Children.Add(italic);
        flags.Children.Add(underline);
        flags.Children.Add(strike);
        fields.Children.Add(flags);
        TextBox foreground = Field(Hex(active.ForegroundColor));
        AddField(Get(S.TextColorHex), foreground);
        TextBox background = Field(Hex(active.BackgroundColor));
        AddField(Get(S.FillColorHex), background);
        TextBlock error = new() { Foreground = Brushes.Firebrick, IsVisible = false };
        fields.Children.Add(error);
        Button save = new() { Content = Get(S.Apply), IsDefault = true, MinWidth = 80 };
        Button cancel = new() { Content = Get(S.Cancel), IsCancel = true, MinWidth = 80 };
        StackPanel buttons = new() { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Spacing = 8 };
        buttons.Children.Add(cancel);
        buttons.Children.Add(save);
        fields.Children.Add(buttons);
        ScrollViewer scroller = new() { Content = fields, Margin = new Thickness(16) };
        Window dialog = new()
        {
            Title = Get(S.FormatCells),
            Width = 500,
            Height = 570,
            CanResize = true,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Content = scroller,
        };
        cancel.Click += (_, _) => dialog.Close();
        save.Click += (_, _) =>
        {
            if (!double.TryParse(fontSize.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double points) ||
                !double.IsFinite(points) || points < 1 || points > 200)
            {
                error.Text = Get(S.FontSizeInvalid);
                error.IsVisible = true;
                return;
            }
            Color? textColor;
            Color? fillColor;
            try
            {
                textColor = string.IsNullOrWhiteSpace(foreground.Text) ? null : Color.Parse(foreground.Text.Trim());
                fillColor = string.IsNullOrWhiteSpace(background.Text) ? null : Color.Parse(background.Text.Trim());
            }
            catch (FormatException)
            {
                error.Text = Get(S.ColorInvalid);
                error.IsVisible = true;
                return;
            }

            sheet.Checkpoint();
            this.applyingCellFormat = true;
            try
            {
                foreach (CellViewModel cell in this.SelectionCells())
                {
                    cell.NumberFormat = string.IsNullOrWhiteSpace(numberFormat.Text) ? null : numberFormat.Text;
                    cell.TextAlignment = alignment.SelectedIndex switch
                    {
                        1 => TextAlignment.Center,
                        2 => TextAlignment.Right,
                        _ => TextAlignment.Left,
                    };
                    cell.FontFamily = string.IsNullOrWhiteSpace(fontName.Text) ? null : fontName.Text.Trim();
                    cell.FontSize = points * 4.0 / 3.0;
                    cell.SetWholeBold(bold.IsChecked == true);
                    cell.SetWholeItalic(italic.IsChecked == true);
                    cell.SetWholeUnderline(underline.IsChecked == true);
                    cell.SetWholeStrikethrough(strike.IsChecked == true);
                    cell.SetWholeForeground(textColor);
                    cell.BackgroundColor = fillColor;
                }
            }
            finally
            {
                this.applyingCellFormat = false;
            }
            this.Rebuild();
            sheet.NotifyMetadataChanged();
            dialog.Close();
        };
        await dialog.ShowDialog(owner);
    }

    private void OpenHyperlink(string target)
    {
        if (target.StartsWith('#'))
        {
            string address = target[1..];
            int separator = address.LastIndexOf('!');
            if (separator >= 0)
            {
                string name = address[..separator].Trim('\'');
                address = address[(separator + 1)..];
                Match crossSheetCell = Regex.Match(address, @"^\$?([A-Za-z]+)\$?(\d+)$");
                if (crossSheetCell.Success && int.TryParse(crossSheetCell.Groups[2].Value, out int crossSheetRow) &&
                    !string.Equals(name, this.sheet?.Name, StringComparison.OrdinalIgnoreCase))
                {
                    this.WorksheetLinkRequested?.Invoke(name, new CellAddress(
                        crossSheetRow - 1, CellAddress.LettersToColumn(crossSheetCell.Groups[1].Value)));
                    return;
                }
            }
            Match match = Regex.Match(address, @"^\$?([A-Za-z]+)\$?(\d+)$");
            if (match.Success && int.TryParse(match.Groups[2].Value, out int row))
            {
                CellAddress cell = new(row - 1, CellAddress.LettersToColumn(match.Groups[1].Value));
                if (this.sheet?.IsValid(cell) == true)
                {
                    this.SelectCell(cell);
                }
            }
            return;
        }
        if (!IsValidLinkTarget(target))
        {
            return;
        }
        try
        {
            Process.Start(new ProcessStartInfo(target) { UseShellExecute = true });
        }
        catch
        {
            // The OS may have no browser/mail handler; the link remains available to copy/edit.
        }
    }

    private MenuFlyout BuildCellMenu()
    {
        MenuFlyout menu = new();
        void Add(string header, Action act)
        {
            MenuItem item = new() { Header = header };
            item.Click += (_, _) => act();
            menu.Items.Add(item);
        }
        Add(Get(S.Cut), this.CutSelection);
        Add(Get(S.Copy), this.CopySelection);
        Add(Get(S.Paste), this.PasteSelection);
        menu.Items.Add(new Separator());
        Add(Get(S.ClearContents), this.ClearSelection);
        menu.Items.Add(new Separator());
        Add(Get(S.InsertComment), () => _ = this.EditCommentAsync());
        Add(Get(S.FormatCellsMenu), () => _ = this.FormatCellsAsync());
        Add(Get(S.HyperlinkMenu), () => _ = this.EditHyperlinkAsync());
        menu.Items.Add(new Separator());
        Add(Get(S.InsertRowAbove), () => this.InsertSelectedRows(this.SelectionRange.Row, this.SelectionRange.RowSpan));
        Add(Get(S.DeleteRow), () => this.DeleteSelectedRows(this.SelectionRange.Row, this.SelectionRange.RowSpan));
        menu.Items.Add(new Separator());
        Add(Get(S.FilterColumnToValue), () =>
        {
            if (this.sheet is { } sheet)
            {
                CellAddress address = sheet.ActiveCell;
                this.FilterColumn(address.Column, sheet.GetCell(address)?.DisplayText ?? string.Empty, sheet.LikelyHeaderRow());
            }
        });
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
