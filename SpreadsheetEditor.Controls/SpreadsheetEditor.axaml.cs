namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;

using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Localization;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.Printing;
using global::SpreadsheetEditor.Controls.ViewModels;
using global::SpreadsheetEditor.Controls.Xlsx;
using S = global::SpreadsheetEditor.Controls.Localization.EditorString;
using static global::SpreadsheetEditor.Controls.Localization.EditorText;

// A reusable spreadsheet editor built on a custom SpreadsheetGrid (not the Avalonia
// DataGrid) so cells can span (merge) and host rich content. The grid renders each cell's text via
// inline runs (bold/italic/underline/colors); clicking a cell sets the active cell. The formula bar
// is a plain TextBox bound live to the active cell, and the toolbar applies whole-cell formatting.
public partial class SpreadsheetEditor : UserControl
{
    public static readonly StyledProperty<int> InitialRowCountProperty =
        AvaloniaProperty.Register<SpreadsheetEditor, int>(nameof(InitialRowCount), 50);

    public static readonly StyledProperty<int> InitialColumnCountProperty =
        AvaloniaProperty.Register<SpreadsheetEditor, int>(nameof(InitialColumnCount), 26);

    public static readonly StyledProperty<bool> IsReadOnlyProperty =
        AvaloniaProperty.Register<SpreadsheetEditor, bool>(nameof(IsReadOnly));

    private static readonly Color[] Palette =
    {
        Colors.Black, Colors.DimGray, Colors.Gray, Colors.White,
        Colors.Red, Colors.Orange, Colors.Gold, Colors.Yellow,
        Colors.Green, Colors.Teal, Colors.Blue, Colors.Indigo,
        Colors.Purple, Colors.HotPink, Colors.Brown, Colors.Transparent,
    };

    private const string DefaultFontName = "Calibri";

    private static readonly string[] FontNames =
    {
        "Calibri", "Cambria", "Arial", "Times New Roman", "Verdana", "Georgia",
        "Tahoma", "Segoe UI", "Courier New", "Consolas", "Comic Sans MS",
    };

    private static readonly int[] FontSizes = { 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72 };

    // Engine-supported functions grouped for the function library (name + signature).
    private static readonly (EditorString Category, (string Name, EditorString Signature)[] Functions)[] FunctionLibrary =
    {
        (S.FunctionMathTrig, new[]
        {
            ("SUM", S.FnSum), ("ABS", S.FnAbs), ("SQRT", S.FnSqrt),
            ("POWER", S.FnPower), ("MOD", S.FnMod), ("INT", S.FnInt),
            ("ROUND", S.FnRound), ("ROUNDUP", S.FnRoundUp), ("ROUNDDOWN", S.FnRoundDown),
        }),
        (S.FunctionStatistical, new[]
        {
            ("AVERAGE", S.FnAverage), ("MIN", S.FnMin), ("MAX", S.FnMax),
            ("COUNT", S.FnCount), ("COUNTA", S.FnCountA),
        }),
        (S.FunctionLogical, new[]
        {
            ("IF", S.FnIf), ("IFERROR", S.FnIfError),
            ("AND", S.FnAnd), ("OR", S.FnOr), ("NOT", S.FnNot),
        }),
        (S.FunctionText, new[]
        {
            ("CONCAT", S.FnConcat), ("CONCATENATE", S.FnConcatenate), ("LEN", S.FnLen),
            ("LEFT", S.FnLeft), ("RIGHT", S.FnRight), ("MID", S.FnMid),
            ("UPPER", S.FnUpper), ("LOWER", S.FnLower), ("TRIM", S.FnTrim),
        }),
        (S.FunctionDateTime, new[]
        {
            ("TODAY", S.FnToday), ("NOW", S.FnNow), ("YEAR", S.FnYear),
            ("MONTH", S.FnMonth), ("DAY", S.FnDay), ("DATE", S.FnDate),
        }),
    };

    // The workbook: an ordered list of sheets, each its own view model; vm points at the active one.
    private readonly List<SpreadsheetViewModel> sheets = new();
    private readonly HashSet<SpreadsheetViewModel> loadingSheets = new();
    private readonly Dictionary<SpreadsheetViewModel, int> stagedSheetIndexes = new();
    private readonly HashSet<int> previewedSheetIndexes = new();
    private readonly HashSet<int> activeSheetLoadIndexes = new();
    private readonly HashSet<int> failedSheetLoadIndexes = new();
    private readonly Dictionary<SpreadsheetViewModel, HashSet<string>> sheetDependencies = new();
    private readonly HashSet<SpreadsheetViewModel> pendingUserRecalculations = new();
    private readonly HashSet<SpreadsheetViewModel> deferredEditRecalculations = new();
    private readonly HashSet<SpreadsheetViewModel> editDirtyNotifications = new();
    private readonly Dictionary<SpreadsheetViewModel, CancellationTokenSource> editRecalculationDelays = new();

    private CancellationTokenSource? workbookLoadCancellation;
    private byte[]? stagedWorkbookBytes;
    private IReadOnlyList<XlsxSheetPreview>? stagedSheetMetadata;
    private XlsxWorkbookReader? stagedWorkbookReader;
    private int? pendingSheetActivationIndex;
    private (string SheetName, CellAddress Address)? pendingWorksheetLink;
    private int workbookLoadGeneration;
    private bool recalculatingWorkbookDependencies;

    private SpreadsheetViewModel vm = new();

    private CellViewModel? activeCell;

    private TextBox? formulaBar;
    private SpreadsheetViewModel? formulaBarEditSheet;

    private bool suppressToolbar;

    private bool initialFocusDone;

    // The sheet whose UndoStateChanged/DataChanged we're currently subscribed to (for CanUndo + ContentChanged).
    private SpreadsheetViewModel? hookedSheet;

    public SpreadsheetEditor()
    {
        this.InitializeComponent();

        this.vm.Initialize(this.InitialRowCount, this.InitialColumnCount);
        this.vm.Name = Get(S.Sheet) + "1";
        this.sheets.Add(this.vm);
        this.ConfigureWorkbookFormulaResolution(this.vm);
        this.DataContext = this.vm;

        this.Grid.Sheet = this.vm;
        this.Grid.CellActivated += address => this.SetActiveCell(address);
        this.Grid.WorksheetLinkRequested += this.NavigateWorksheetLink;
        this.Grid.EditCompleted += this.CompleteDeferredEditRecalculation;

        // Keep spreadsheet shortcuts inside the editor. In particular, Ctrl+F must not bubble to a
        // host application's global search box.
        this.AddHandler(KeyDownEvent, this.OnEditorKeyDown, RoutingStrategies.Tunnel);
        this.FindNextButton.Click += (_, _) => this.FindMatch(backwards: false);
        this.FindPreviousButton.Click += (_, _) => this.FindMatch(backwards: true);
        this.ReplaceButton.Click += (_, _) => this.ReplaceCurrentMatch();
        this.ReplaceAllButton.Click += (_, _) => this.ReplaceAllMatches();
        this.CloseFindButton.Click += (_, _) => this.CloseFindReplace();
        this.FindTextBox.TextChanged += (_, _) => this.FindStatusText.Text = string.Empty;
        this.FindTextBox.KeyDown += (_, e) => this.OnFindBoxKeyDown(e);
        this.ReplaceTextBox.KeyDown += (_, e) => this.OnFindBoxKeyDown(e);

        this.AddSheetButton.Click += (_, _) => this.AddSheet();

        // Zoom control (bottom bar) — buttons + a live "100%" readout.
        this.ZoomInButton.Click += (_, _) => this.Grid.ZoomIn();
        this.ZoomOutButton.Click += (_, _) => this.Grid.ZoomOut();
        this.ZoomResetButton.Click += (_, _) => this.Grid.ResetZoom();
        this.Grid.ZoomChanged += zoom => this.ZoomResetButton.Content = $"{Math.Round(zoom * 100)}%";

        this.CreateFormulaBar();

        this.BuildPalette(this.ForegroundPalette, this.ApplyForeground);
        this.BuildPalette(this.BackgroundPalette, this.ApplyBackground);

        // B/I/U + text color route to the in-cell editor (per-character when the rich editor is
        // active; whole-cell otherwise). Guarded so SyncToolbar setting IsChecked doesn't re-toggle.
        this.BoldButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar) { this.Grid.ToggleBold(); this.SyncFormatButtons(); } };
        this.ItalicButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar) { this.Grid.ToggleItalic(); this.SyncFormatButtons(); } };
        this.UnderlineButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar) { this.Grid.ToggleUnderline(); this.SyncFormatButtons(); } };
        this.StrikethroughButton.IsCheckedChanged += (_, _) => { if (!this.suppressToolbar) { this.Grid.ToggleStrikethrough(); this.SyncFormatButtons(); } };

        // Keep B/I/U reflecting the caret/selection while editing.
        this.Grid.SelectionFormatChanged += (_, _) => this.SyncFormatButtons();

        this.SetupFontControls();
        this.AlignLeftButton.IsCheckedChanged += (_, _) =>
        {
            if (!this.suppressToolbar && this.AlignLeftButton.IsChecked == true)
            {
                this.SetAlignment(TextAlignment.Left);
            }
        };
        this.AlignCenterButton.IsCheckedChanged += (_, _) =>
        {
            if (!this.suppressToolbar && this.AlignCenterButton.IsChecked == true)
            {
                this.SetAlignment(TextAlignment.Center);
            }
        };
        this.AlignRightButton.IsCheckedChanged += (_, _) =>
        {
            if (!this.suppressToolbar && this.AlignRightButton.IsChecked == true)
            {
                this.SetAlignment(TextAlignment.Right);
            }
        };
        this.Grid.ChartEditRequested += async chart => await this.EditChartAsync(chart);

        // History + clipboard buttons.
        this.UndoButton.Click += (_, _) => this.Undo();
        this.RedoButton.Click += (_, _) => this.Redo();
        this.CutButton.Click += (_, _) => this.Grid.CutSelection();
        this.CopyButton.Click += (_, _) => this.Grid.CopySelection();
        this.PasteButton.Click += (_, _) => this.Grid.PasteSelection();
        this.PrintButton.Click += (_, _) => this.PrintPdf();

        // Spreadsheet-specific actions live in the overflow menu.
        this.OverflowButton.Flyout = this.BuildOverflowMenu();
        this.SymbolButton.Flyout = this.BuildSymbolFlyout();

        this.SetActiveCell(new CellAddress(0, 0));
        this.RebuildTabs();
        this.HookSheet();
        this.ApplyReadOnly();

        // Recolor the sheet tabs when the light/dark theme changes at runtime.
        this.ActualThemeVariantChanged += (_, _) => this.RebuildTabs();

        // Start with A1 selected (selection mode), not the auto-focused formula bar or an open editor.
        this.Loaded += (_, _) =>
        {
            this.RebuildTabs(); // theme is resolved by now — recolor tabs for the actual variant

            if (this.initialFocusDone)
            {
                return;
            }

            this.initialFocusDone = true;
            Dispatcher.UIThread.Post(this.Grid.FocusForSelection, DispatcherPriority.Background);
        };
    }

    // The current (active) sheet data (rendering-agnostic). Host apps read/write this.
    public SpreadsheetViewModel ViewModel => this.vm;

    // All sheets in the workbook, in tab order.
    public IReadOnlyList<SpreadsheetViewModel> Sheets => this.sheets;

    public int InitialRowCount
    {
        get => this.GetValue(InitialRowCountProperty);
        set => this.SetValue(InitialRowCountProperty, value);
    }

    public int InitialColumnCount
    {
        get => this.GetValue(InitialColumnCountProperty);
        set => this.SetValue(InitialColumnCountProperty, value);
    }

    // When true, the editor is view-only: no cell editing, typing, paste/cut/clear, or formatting.
    // Mirrors RichTextEditor.IsReadOnly so both controls integrate the same way.
    public bool IsReadOnly
    {
        get => this.GetValue(IsReadOnlyProperty);
        set => this.SetValue(IsReadOnlyProperty, value);
    }

    // The selected cell range (anchor↔focus), mirroring RichTextEditor's selection surface.
    public CellRange SelectedRange => this.Grid.SelectionRange;

    // True when there is a change to undo / redo on the active sheet.
    public bool CanUndo => this.vm.CanUndo;

    public bool CanRedo => this.vm.CanRedo;

    // Raised when an edit is committed (or a non-editor mutation occurs). Hosts commonly save in
    // response, so per-character notifications would serialize a workbook on every keystroke.
    public event EventHandler? ContentChanged;

    // Raised when the active cell / selected range changes — analogous to RichTextEditor.SelectionChanged.
    public event EventHandler? SelectionChanged;

    // Raised when IsReadOnly changes — analogous to RichTextEditor.IsReadOnlyChanged.
    public event EventHandler? IsReadOnlyChanged;

    protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
    {
        base.OnPropertyChanged(change);
        if (change.Property == IsReadOnlyProperty)
        {
            this.ApplyReadOnly();
        }
    }

    private void ApplyReadOnly(bool raiseEvent = true)
    {
        bool readOnly = this.IsReadOnly || this.loadingSheets.Contains(this.vm);
        this.Grid.ReadOnly = readOnly;
        this.ReplaceButton.IsEnabled = !readOnly;
        this.ReplaceAllButton.IsEnabled = !readOnly;
        if (this.formulaBar is not null)
        {
            this.formulaBar.IsReadOnly = readOnly;
        }

        if (raiseEvent)
        {
            this.IsReadOnlyChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    // Index of the active sheet in sheets.
    private int ActiveSheetIndex => this.sheets.IndexOf(this.vm);

    // ----- Public API (for host apps) -----

    public void NewSheet(int rows, int columns)
    {
        this.CancelWorkbookLoad();
        this.DetachActiveCell();
        this.vm = new SpreadsheetViewModel { Name = Get(S.Sheet) + "1" };
        this.vm.Initialize(rows, columns);
        this.sheets.Clear();
        this.sheets.Add(this.vm);
        this.sheetDependencies.Clear();
        this.pendingUserRecalculations.Clear();
        this.ConfigureWorkbookFormulaResolution(this.vm);
        this.DataContext = this.vm;
        this.Grid.Sheet = this.vm;
        this.SetActiveCell(new CellAddress(0, 0));
        this.RebuildTabs();
        this.HookSheet();
        this.ApplyReadOnly(raiseEvent: false);
    }

    public void LoadRows(IReadOnlyList<IReadOnlyList<string>> data)
    {
        this.DetachActiveCell();
        this.vm.LoadRows(data);
        this.Grid.Rebuild();
        this.SetActiveCell(new CellAddress(0, 0));
    }

    public IReadOnlyList<IReadOnlyList<string>> GetRows() => this.vm.ToRows();

    public void LoadCellData(IReadOnlyList<IReadOnlyList<CellData>> data)
    {
        this.DetachActiveCell();
        this.vm.LoadCellData(data);
        this.Grid.Rebuild();
        this.SetActiveCell(new CellAddress(0, 0));
    }

    public IReadOnlyList<IReadOnlyList<CellData>> GetCellData() => this.vm.ToCellData();

    // ----- Excel (.xlsx) I/O (backed by the bundled ClosedXML + OOXML readers) -----

    // Reads an .xlsx stream — every worksheet, with formulas, formatting, merges, charts, images, data
    // bars, and named ranges — into the editor (one tab per worksheet).
    public void OpenXlsx(Stream stream) => this.LoadWorkbook(XlsxFile.ReadWorkbook(stream));

    // Publishes all tab names/colors plus a screen-sized first-sheet preview. Other worksheets remain
    // lightweight placeholders until selected; selection publishes the same fast preview before its
    // full-fidelity worksheet is prepared away from the UI thread.
    public async Task OpenXlsxAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        this.CancelWorkbookLoad();
        CancellationTokenSource loadCancellation = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        CancellationToken loadToken = loadCancellation.Token;
        this.workbookLoadCancellation = loadCancellation;
        int generation = this.workbookLoadGeneration;
        bool keepLoadLifetime = false;
        this.LoadingOverlay.IsVisible = true;
        try
        {
            // Copy first so the caller can dispose its stream while parsing continues and so even a
            // non-seekable storage-provider stream can be read by ClosedXML on a worker thread.
            using MemoryStream buffer = new();
            await stream.CopyToAsync(buffer, loadToken);
            byte[] bytes = buffer.ToArray();

            // The lightweight preview reads workbook metadata, every tab color, first-sheet cells,
            // and basic styles directly from OOXML. It avoids ClosedXML's eager whole-workbook
            // construction and sizes the preview to roughly one visible viewport.
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.Render);
            (int previewRows, int previewColumns) = this.Grid.SuggestedPreviewSize();
            (XlsxWorkbookPreview Preview, List<SpreadsheetViewModel> Staged) prepared = await Task.Run(
                () =>
                {
                    loadToken.ThrowIfCancellationRequested();
                    XlsxWorkbookPreview preview = XlsxFile.ReadPreview(
                        bytes,
                        loadToken,
                        previewRows,
                        previewColumns);
                    return (preview, PreparePreviewWorkbook(preview));
                },
                loadToken);
            loadToken.ThrowIfCancellationRequested();
            if (generation != this.workbookLoadGeneration)
            {
                return;
            }

            this.stagedWorkbookBytes = bytes;
            this.stagedSheetMetadata = prepared.Preview.Sheets;
            this.stagedWorkbookReader = new XlsxWorkbookReader(bytes);
            this.LoadStagedWorkbook(prepared.Staged);
            this.LoadingOverlay.IsVisible = false;
            keepLoadLifetime = true;

            // The first worksheet is selected as part of opening the workbook. Its preview is already
            // visible, so begin only that worksheet's full import. Every other worksheet waits for a
            // tab click before doing any cell work.
            this.StartStagedSheetLoad(prepared.Staged[0], generation, loadToken);
        }
        catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
        {
            // A newer workbook/load operation superseded this one.
        }
        catch (OperationCanceledException)
        {
            this.StopStagedLoading(generation, Get(S.WorksheetLoadCanceled));
            throw;
        }
        catch
        {
            this.StopStagedLoading(generation, Get(S.WorksheetLoadFailed));
            throw;
        }
        finally
        {
            if (generation == this.workbookLoadGeneration)
            {
                this.LoadingOverlay.IsVisible = false;
            }
            if (!keepLoadLifetime && ReferenceEquals(this.workbookLoadCancellation, loadCancellation))
            {
                this.workbookLoadCancellation = null;
                loadCancellation.Dispose();
            }
        }
    }

    private sealed record WorkbookSaveSnapshot(
        List<XlsxSheetWriteData?> Sheets,
        List<(int OutputIndex, int SourceIndex, string Name)> PendingSheets,
        byte[]? SourceBytes);

    // Writes every worksheet, including lazily loaded sheets that are still placeholders. Never
    // serialize a preview/placeholder as the real worksheet: it would permanently discard data.
    public void SaveXlsx(Stream stream) => WriteWorkbookSnapshot(stream, this.CaptureWorkbookForSave(), CancellationToken.None);

    // Hosts that save large workbooks should use this overload: only the UI-owned snapshot is taken
    // on the UI thread; source-sheet materialization and XLSX serialization run in the background.
    public async Task SaveXlsxAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        WorkbookSaveSnapshot snapshot = Dispatcher.UIThread.CheckAccess()
            ? this.CaptureWorkbookForSave()
            : await Dispatcher.UIThread.InvokeAsync(this.CaptureWorkbookForSave, DispatcherPriority.Normal);
        await Task.Run(() => WriteWorkbookSnapshot(stream, snapshot, cancellationToken), cancellationToken)
            .ConfigureAwait(false);
    }

    private WorkbookSaveSnapshot CaptureWorkbookForSave()
    {
        // Saving is a snapshot operation. Ending the active edit here used to trigger a full grid
        // rebuild on every host auto-save, visibly flashing the worksheet while the user typed.
        // A save consumes the current dirty notification. If the caret stays in the editor and
        // more text arrives afterward, that next change must notify the host again.
        this.editDirtyNotifications.Clear();
        List<XlsxSheetWriteData?> workbook = new(this.sheets.Count);
        List<(int OutputIndex, int SourceIndex, string Name)> pending = new();
        foreach (SpreadsheetViewModel sheet in this.sheets)
        {
            if (this.loadingSheets.Contains(sheet))
            {
                if (!this.stagedSheetIndexes.TryGetValue(sheet, out int sourceIndex))
                {
                    throw new InvalidOperationException("A pending worksheet has no source workbook index.");
                }
                pending.Add((workbook.Count, sourceIndex, sheet.Name));
                workbook.Add(null);
                continue;
            }

            workbook.Add(new XlsxSheetWriteData(
                sheet.Name,
                CaptureSheetData(sheet),
                SpreadsheetGrid.BuildChartWriteData(sheet, sheet.Name)));
        }

        if (pending.Count > 0 && this.stagedWorkbookBytes is null)
        {
            throw new InvalidOperationException("Pending worksheets cannot be saved without the original workbook.");
        }
        return new WorkbookSaveSnapshot(workbook, pending, this.stagedWorkbookBytes);
    }

    private static void WriteWorkbookSnapshot(
        Stream stream,
        WorkbookSaveSnapshot snapshot,
        CancellationToken cancellationToken)
    {
        if (snapshot.PendingSheets.Count > 0)
        {
            Dictionary<int, SheetData> sourceSheets = XlsxFile.ReadSelectedWorksheets(
                snapshot.SourceBytes!,
                snapshot.PendingSheets.Select(static item => item.SourceIndex),
                cancellationToken);
            foreach ((int outputIndex, int sourceIndex, string name) in snapshot.PendingSheets)
            {
                SheetData data = sourceSheets[sourceIndex];
                SpreadsheetViewModel sheet = PrepareSheet(name, data);
                snapshot.Sheets[outputIndex] = new XlsxSheetWriteData(
                    name, data, SpreadsheetGrid.BuildChartWriteData(sheet, name));
            }
        }

        cancellationToken.ThrowIfCancellationRequested();
        XlsxFile.WriteWorkbook(stream, snapshot.Sheets.Select(static sheet => sheet!).ToList());
    }

    private static SheetData CaptureSheetData(SpreadsheetViewModel sheet) => new()
    {
        Cells = sheet.ToCellData(),
        ColumnWidths = sheet.Columns.Select(static column => column.Width).ToList(),
        Merges = sheet.Merges.ToArray(),
        Images = sheet.Images.ToArray(),
        Charts = sheet.Charts.ToArray(),
        Names = new Dictionary<string, string>(sheet.Names),
        DataBars = sheet.DataBars.ToArray(),
        CellValueRules = sheet.CellValueRules.ToArray(),
        AutoFilterRange = sheet.AutoFilterRange,
        HiddenRows = sheet.Rows.Select((row, index) => (row, index))
            .Where(static item => item.row.IsHidden).Select(static item => item.index).ToArray(),
        TabColor = sheet.TabColor,
    };

    // Loads a full single-sheet snapshot (resets the workbook to one sheet): cells, per-column
    // widths, merges, images, charts, data bars, and names.
    public void LoadSheet(SheetData sheet) =>
        this.LoadWorkbook(new[] { (Get(S.Sheet) + "1", sheet) });

    // Replaces the whole workbook with the given named sheets (e.g. every worksheet of an
    // opened .xlsx) and opens the first one.
    public void LoadWorkbook(IReadOnlyList<(string Name, SheetData Data)> sheets)
    {
        this.CancelWorkbookLoad();
        if (sheets.Count == 0)
        {
            this.NewSheet(this.InitialRowCount, this.InitialColumnCount);
            return;
        }

        this.LoadPreparedWorkbook(PrepareWorkbook(sheets));
    }

    private static List<SpreadsheetViewModel> PrepareWorkbook(IReadOnlyList<(string Name, SheetData Data)> sheets)
    {
        HashSet<string> names = new(StringComparer.OrdinalIgnoreCase);
        List<SpreadsheetViewModel> prepared = new(sheets.Count);
        foreach ((string name, SheetData data) in sheets)
        {
            string desired = string.IsNullOrWhiteSpace(name) ? Get(S.Sheet) : name.Trim();
            string unique = desired;
            for (int suffix = 2; !names.Add(unique); suffix++)
            {
                unique = $"{desired}{suffix}";
            }

            prepared.Add(PrepareSheet(unique, data));
        }

        return prepared;
    }

    private void LoadPreparedWorkbook(IReadOnlyList<SpreadsheetViewModel> prepared)
    {
        this.DetachActiveCell();
        this.loadingSheets.Clear();
        this.sheetDependencies.Clear();
        this.pendingUserRecalculations.Clear();
        this.sheets.Clear();
        foreach (SpreadsheetViewModel sheet in prepared)
        {
            this.sheets.Add(sheet);
            this.ConfigureWorkbookFormulaResolution(sheet);
            this.sheetDependencies[sheet] = new HashSet<string>(sheet.ReferencedSheetNames(), StringComparer.OrdinalIgnoreCase);
        }

        this.ActivateSheet(0);
        this.RebuildTabs();
    }

    private static SpreadsheetViewModel PrepareSheet(string name, SheetData data)
    {
        SpreadsheetViewModel sheet = new() { Name = name };
        ApplySheetData(sheet, data);
        return sheet;
    }

    private static List<SpreadsheetViewModel> PreparePreviewWorkbook(XlsxWorkbookPreview preview)
    {
        HashSet<string> names = new(StringComparer.OrdinalIgnoreCase);
        List<SpreadsheetViewModel> staged = new(preview.Sheets.Count);
        for (int index = 0; index < preview.Sheets.Count; index++)
        {
            XlsxSheetPreview metadata = preview.Sheets[index];
            string desired = string.IsNullOrWhiteSpace(metadata.Name) ? Get(S.Sheet) : metadata.Name.Trim();
            string unique = desired;
            for (int suffix = 2; !names.Add(unique); suffix++)
            {
                unique = $"{desired}{suffix}";
            }

            if (index == 0)
            {
                staged.Add(PrepareSheet(unique, preview.FirstSheet));
                continue;
            }

            SpreadsheetViewModel placeholder = new() { Name = unique };
            placeholder.Initialize(1, 1);
            placeholder.GetCell(new CellAddress(0, 0))!.Text = Get(S.WorksheetLoadPrompt);
            placeholder.TabColor = metadata.TabColor;
            staged.Add(placeholder);
        }

        return staged;
    }

    private void LoadStagedWorkbook(IReadOnlyList<SpreadsheetViewModel> staged)
    {
        this.DetachActiveCell();
        this.sheets.Clear();
        this.loadingSheets.Clear();
        this.stagedSheetIndexes.Clear();
        this.previewedSheetIndexes.Clear();
        this.activeSheetLoadIndexes.Clear();
        this.failedSheetLoadIndexes.Clear();
        this.sheetDependencies.Clear();
        this.pendingUserRecalculations.Clear();
        this.pendingSheetActivationIndex = null;
        for (int index = 0; index < staged.Count; index++)
        {
            SpreadsheetViewModel sheet = staged[index];
            this.sheets.Add(sheet);
            this.loadingSheets.Add(sheet);
            this.stagedSheetIndexes[sheet] = index;
            this.ConfigureWorkbookFormulaResolution(sheet);
        }
        this.previewedSheetIndexes.Add(0);

        this.ActivateSheet(0);
        this.RebuildTabs();
    }

    private void StartStagedSheetLoad(
        SpreadsheetViewModel sheet,
        int generation,
        CancellationToken cancellationToken,
        bool requireActive = true)
    {
        if (generation != this.workbookLoadGeneration ||
            !this.loadingSheets.Contains(sheet) ||
            !this.stagedSheetIndexes.TryGetValue(sheet, out int sheetIndex) ||
            !this.activeSheetLoadIndexes.Add(sheetIndex))
        {
            return;
        }

        this.failedSheetLoadIndexes.Remove(sheetIndex);
        (int previewRows, int previewColumns) = this.Grid.SuggestedPreviewSize();
        this.RebuildTabs();
        _ = this.LoadStagedSheetAsync(
            sheet,
            sheetIndex,
            previewRows,
            previewColumns,
            generation,
            cancellationToken,
            requireActive);
    }

    private async Task LoadStagedSheetAsync(
        SpreadsheetViewModel target,
        int sheetIndex,
        int previewRows,
        int previewColumns,
        int generation,
        CancellationToken cancellationToken,
        bool requireActive)
    {
        SpreadsheetViewModel current = target;
        try
        {
            byte[] bytes = this.stagedWorkbookBytes ?? throw new InvalidOperationException(Get(S.StagedWorkbookMissing));
            IReadOnlyList<XlsxSheetPreview> metadata = this.stagedSheetMetadata ??
                throw new InvalidOperationException(Get(S.StagedMetadataMissing));
            XlsxWorkbookReader reader = this.stagedWorkbookReader ??
                throw new InvalidOperationException(Get(S.StagedReaderMissing));
            if (sheetIndex < 0 || sheetIndex >= metadata.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex));
            }

            if (!this.previewedSheetIndexes.Contains(sheetIndex))
            {
                string previewName = current.Name;
                SpreadsheetViewModel preview = await Task.Run(
                    () => PrepareSheet(
                        previewName,
                        XlsxFile.ReadWorksheetPreview(
                            bytes,
                            metadata[sheetIndex],
                            previewRows,
                            previewColumns,
                            cancellationToken)),
                    cancellationToken);
                SpreadsheetViewModel? replacement = await Dispatcher.UIThread.InvokeAsync(
                    () => this.ReplaceStagedSheet(generation, sheetIndex, current, preview, completed: false),
                    DispatcherPriority.Background);
                if (replacement is null)
                {
                    return;
                }
                current = replacement;
            }

            // A brief stability window prevents rapid tab browsing from queuing a full import for
            // every tab the pointer passes through. Formula dependencies bypass this check.
            if (requireActive)
            {
                await Task.Delay(150, cancellationToken).ConfigureAwait(false);
                bool stillActive = await Dispatcher.UIThread.InvokeAsync(
                    () => this.stagedSheetIndexes.TryGetValue(this.vm, out int activeIndex) && activeIndex == sheetIndex,
                    DispatcherPriority.Background);
                if (!stillActive)
                {
                    return;
                }
            }

            SheetData fullData = await reader.ReadWorksheetAsync(sheetIndex, cancellationToken).ConfigureAwait(false);
            string loadedName = current.Name;
            (SpreadsheetViewModel Loaded, HashSet<string> Dependencies) completedSheet = await Task.Run(
                () =>
                {
                    SpreadsheetViewModel loaded = PrepareSheet(loadedName, fullData);
                    return (loaded, new HashSet<string>(loaded.ReferencedSheetNames(), StringComparer.OrdinalIgnoreCase));
                },
                cancellationToken).ConfigureAwait(false);

            await Dispatcher.UIThread.InvokeAsync(
                () => this.ReplaceStagedSheet(
                    generation,
                    sheetIndex,
                    current,
                    completedSheet.Loaded,
                    completed: true,
                    completedSheet.Dependencies,
                    cancellationToken),
                DispatcherPriority.Background);
        }
        catch (OperationCanceledException)
        {
            await Dispatcher.UIThread.InvokeAsync(
                () => this.MarkStagedSheetLoadFailed(generation, sheetIndex, Get(S.WorksheetLoadCanceled)),
                DispatcherPriority.Background);
        }
        catch
        {
            await Dispatcher.UIThread.InvokeAsync(
                () => this.MarkStagedSheetLoadFailed(generation, sheetIndex, Get(S.WorksheetLoadFailed)),
                DispatcherPriority.Background);
        }
        finally
        {
            await Dispatcher.UIThread.InvokeAsync(
                () =>
                {
                    if (generation == this.workbookLoadGeneration)
                    {
                        this.activeSheetLoadIndexes.Remove(sheetIndex);
                        this.RebuildTabs();
                    }
                },
                DispatcherPriority.Background);
        }
    }

    private SpreadsheetViewModel? ReplaceStagedSheet(
        int generation,
        int sheetIndex,
        SpreadsheetViewModel target,
        SpreadsheetViewModel replacement,
        bool completed,
        IReadOnlySet<string>? dependencies = null,
        CancellationToken cancellationToken = default)
    {
        if (generation != this.workbookLoadGeneration ||
            !this.stagedSheetIndexes.TryGetValue(target, out int targetIndex) ||
            targetIndex != sheetIndex)
        {
            return null;
        }

        int index = this.sheets.IndexOf(target);
        if (index < 0)
        {
            return null;
        }

        replacement.Name = target.Name;
        replacement.TabColor ??= target.TabColor;
        this.ConfigureWorkbookFormulaResolution(replacement);
        bool active = ReferenceEquals(target, this.vm);
        bool activatePreview = !completed && this.pendingSheetActivationIndex == sheetIndex;
        CellAddress activeCell = target.ActiveCell;
        this.sheets[index] = replacement;
        this.stagedSheetIndexes.Remove(target);
        this.stagedSheetIndexes[replacement] = sheetIndex;
        this.sheetDependencies.Remove(target);
        if (this.pendingUserRecalculations.Remove(target))
        {
            this.pendingUserRecalculations.Add(replacement);
        }
        this.loadingSheets.Remove(target);
        if (completed)
        {
            this.failedSheetLoadIndexes.Remove(sheetIndex);
            this.sheetDependencies[replacement] = new HashSet<string>(
                dependencies ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase),
                StringComparer.OrdinalIgnoreCase);
        }
        else
        {
            this.loadingSheets.Add(replacement);
            this.previewedSheetIndexes.Add(sheetIndex);
        }

        if (active || activatePreview)
        {
            if (activatePreview)
            {
                this.pendingSheetActivationIndex = null;
            }
            this.DetachActiveCell();
            this.vm = replacement;
            this.DataContext = replacement;
            this.Grid.Sheet = replacement;
            this.SetActiveCell(new CellAddress(
                Math.Clamp(activeCell.Row, 0, replacement.RowCount - 1),
                Math.Clamp(activeCell.Column, 0, replacement.ColumnCount - 1)));
            this.HookSheet();
            this.ApplyReadOnly(raiseEvent: false);
            this.CompletePendingWorksheetLink(completed);
        }

        this.RebuildTabs();
        if (completed)
        {
            this.TryRunPendingUserRecalculations();
        }
        if (completed && this.loadingSheets.Count == 0)
        {
            this.stagedWorkbookBytes = null;
            this.stagedSheetMetadata = null;
            this.stagedWorkbookReader?.DisposeWhenIdle();
            this.stagedWorkbookReader = null;
            this.workbookLoadCancellation?.Dispose();
            this.workbookLoadCancellation = null;
        }
        return replacement;
    }

    private void ConfigureWorkbookFormulaResolution(SpreadsheetViewModel sheet)
    {
        sheet.SetWorkbookCellResolver(this.ResolveWorkbookCell);
        sheet.SetEditRecalculationGate(() => this.RequestUserEditRecalculation(sheet));
    }

    private bool RequestUserEditRecalculation(SpreadsheetViewModel sheet)
    {
        this.CancelScheduledEditRecalculation(sheet);
        bool editing = this.IsActivelyEditing(sheet);
        if (editing)
        {
            this.deferredEditRecalculations.Add(sheet);
            // Dirty tracking must not wait for focus to leave the cell: hosts may save or close
            // while the editor still has the caret. Notify once per edit, not on every keystroke.
            if (this.editDirtyNotifications.Add(sheet))
            {
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }
        }
        else if (this.AreDependenciesReady(sheet))
        {
            CancellationTokenSource delay = new();
            this.editRecalculationDelays[sheet] = delay;
            _ = this.RecalculateAfterTypingPauseAsync(sheet, delay, this.workbookLoadGeneration);
        }
        else
        {
            this.QueuePendingUserRecalculation(sheet);
        }

        // Edits outside an active cell session notify immediately. Active sessions notify once
        // above and again after completion, so hosts never need to save on every keystroke.
        if (!editing)
        {
            this.ContentChanged?.Invoke(this, EventArgs.Empty);
        }
        return false;
    }

    private bool IsActivelyEditing(SpreadsheetViewModel sheet) =>
        ReferenceEquals(sheet, this.vm) &&
        (this.Grid.IsEditing ||
         (ReferenceEquals(this.formulaBarEditSheet, sheet) && this.formulaBar?.IsFocused == true));

    private void CompleteDeferredEditRecalculation(SpreadsheetViewModel sheet)
    {
        this.editDirtyNotifications.Remove(sheet);
        if (!this.deferredEditRecalculations.Contains(sheet))
        {
            return;
        }

        // Let the committed cell and selection paint first. If another edit has already started by
        // this dispatcher pass, the next completion event owns the recalculation instead.
        Dispatcher.UIThread.Post(
            () =>
            {
                if (!this.deferredEditRecalculations.Contains(sheet) || this.IsActivelyEditing(sheet))
                {
                    return;
                }

                this.deferredEditRecalculations.Remove(sheet);
                this.CancelScheduledEditRecalculation(sheet);
                if (this.AreDependenciesReady(sheet))
                {
                    this.RecalculateWorkbookFrom(sheet.Name, recalculateSource: true);
                    if (!ReferenceEquals(this.hookedSheet, sheet))
                    {
                        this.ContentChanged?.Invoke(this, EventArgs.Empty);
                    }
                }
                else
                {
                    this.QueuePendingUserRecalculation(sheet);
                    this.ContentChanged?.Invoke(this, EventArgs.Empty);
                }
            },
            DispatcherPriority.Background);
    }

    private void QueuePendingUserRecalculation(SpreadsheetViewModel sheet)
    {
        this.pendingUserRecalculations.Add(sheet);
        if (this.workbookLoadCancellation is { } loadCancellation)
        {
            this.QueueReferencedSheetLoads(sheet, this.workbookLoadGeneration, loadCancellation.Token);
        }
    }

    private async Task RecalculateAfterTypingPauseAsync(
        SpreadsheetViewModel sheet,
        CancellationTokenSource delay,
        int generation)
    {
        try
        {
            await Task.Delay(300, delay.Token).ConfigureAwait(false);
            await Dispatcher.UIThread.InvokeAsync(
                () =>
                {
                    if (generation != this.workbookLoadGeneration ||
                        !this.sheets.Contains(sheet) ||
                        !this.editRecalculationDelays.TryGetValue(sheet, out CancellationTokenSource? current) ||
                        !ReferenceEquals(current, delay))
                    {
                        return;
                    }

                    this.editRecalculationDelays.Remove(sheet);
                    delay.Dispose();
                    this.RecalculateWorkbookFrom(sheet.Name, recalculateSource: true);
                },
                DispatcherPriority.Background);
        }
        catch (OperationCanceledException)
        {
            // A newer keystroke replaced this debounce window.
        }
    }

    private void CancelScheduledEditRecalculation(SpreadsheetViewModel sheet)
    {
        if (!this.editRecalculationDelays.Remove(sheet, out CancellationTokenSource? delay))
        {
            return;
        }

        delay.Cancel();
        delay.Dispose();
    }

    private void CancelScheduledEditRecalculations()
    {
        foreach (CancellationTokenSource delay in this.editRecalculationDelays.Values)
        {
            delay.Cancel();
            delay.Dispose();
        }
        this.editRecalculationDelays.Clear();
    }

    private CellValue? ResolveWorkbookCell(string sheetName, CellAddress address)
    {
        SpreadsheetViewModel? sheet = this.sheets.FirstOrDefault(
            candidate => candidate.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
        return sheet is null || this.loadingSheets.Contains(sheet)
            ? null
            : sheet.EvaluateCell(address);
    }

    private void QueueReferencedSheetLoads(
        SpreadsheetViewModel source,
        int generation,
        CancellationToken cancellationToken)
    {
        if (!this.sheetDependencies.TryGetValue(source, out HashSet<string>? dependencies))
        {
            return;
        }

        foreach (string sheetName in dependencies)
        {
            SpreadsheetViewModel? dependency = this.sheets.FirstOrDefault(
                candidate => candidate.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            if (dependency is not null && this.loadingSheets.Contains(dependency))
            {
                this.StartStagedSheetLoad(
                    dependency,
                    generation,
                    cancellationToken,
                    requireActive: false);
            }
        }
    }

    private void RecalculateWorkbookFrom(
        string changedSheetName,
        bool recalculateSource,
        bool recalculateAll = false)
    {
        if (this.recalculatingWorkbookDependencies)
        {
            return;
        }

        this.recalculatingWorkbookDependencies = true;
        try
        {
            if (recalculateAll)
            {
                foreach (SpreadsheetViewModel sheet in this.sheets.Where(sheet => !this.loadingSheets.Contains(sheet)))
                {
                    sheet.Recalculate();
                }
                return;
            }

            HashSet<SpreadsheetViewModel> recalculated = new();
            Queue<string> changed = new();
            changed.Enqueue(changedSheetName);
            SpreadsheetViewModel? source = this.sheets.FirstOrDefault(
                sheet => sheet.Name.Equals(changedSheetName, StringComparison.OrdinalIgnoreCase));
            if (recalculateSource && source is not null && !this.loadingSheets.Contains(source))
            {
                source.Recalculate();
                recalculated.Add(source);
            }

            while (changed.Count > 0)
            {
                string dependencyName = changed.Dequeue();
                foreach ((SpreadsheetViewModel sheet, HashSet<string> dependencies) in this.sheetDependencies)
                {
                    if (recalculated.Contains(sheet) ||
                        this.loadingSheets.Contains(sheet) ||
                        !dependencies.Contains(dependencyName))
                    {
                        continue;
                    }

                    if (!this.AreDependenciesReady(sheet))
                    {
                        this.pendingUserRecalculations.Add(sheet);
                        if (this.workbookLoadCancellation is { } loadCancellation)
                        {
                            this.QueueReferencedSheetLoads(sheet, this.workbookLoadGeneration, loadCancellation.Token);
                        }
                        continue;
                    }

                    sheet.Recalculate();
                    recalculated.Add(sheet);
                    changed.Enqueue(sheet.Name);
                }
            }
        }
        finally
        {
            this.recalculatingWorkbookDependencies = false;
        }
    }

    private void TryRunPendingUserRecalculations()
    {
        if (this.recalculatingWorkbookDependencies)
        {
            return;
        }

        List<SpreadsheetViewModel> ready = this.pendingUserRecalculations
            .Where(sheet => !this.loadingSheets.Contains(sheet) && this.AreDependenciesReady(sheet))
            .ToList();
        foreach (SpreadsheetViewModel sheet in ready)
        {
            this.pendingUserRecalculations.Remove(sheet);
            this.CancelScheduledEditRecalculation(sheet);
            this.RecalculateWorkbookFrom(sheet.Name, recalculateSource: true);
        }
    }

    private bool AreDependenciesReady(SpreadsheetViewModel sheet)
    {
        if (!this.sheetDependencies.TryGetValue(sheet, out HashSet<string>? dependencies))
        {
            return true;
        }

        return dependencies.All(sheetName =>
        {
            SpreadsheetViewModel? dependency = this.sheets.FirstOrDefault(
                candidate => candidate.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));
            return dependency is not null && !this.loadingSheets.Contains(dependency);
        });
    }

    private void MarkStagedSheetLoadFailed(int generation, int sheetIndex, string placeholderText)
    {
        if (generation != this.workbookLoadGeneration)
        {
            return;
        }

        this.failedSheetLoadIndexes.Add(sheetIndex);
        SpreadsheetViewModel? sheet = this.stagedSheetIndexes
            .FirstOrDefault(pair => pair.Value == sheetIndex).Key;
        CellViewModel? cell = sheet?.GetCell(new CellAddress(0, 0));
        if (cell?.Text == Get(S.WorksheetLoadPrompt))
        {
            cell.Text = placeholderText;
        }
        this.ApplyReadOnly(raiseEvent: false);
        this.RebuildTabs();
    }

    private void CancelWorkbookLoad()
    {
        this.workbookLoadGeneration++;
        this.workbookLoadCancellation?.Cancel();
        this.workbookLoadCancellation?.Dispose();
        this.workbookLoadCancellation = null;
        this.stagedWorkbookBytes = null;
        this.stagedSheetMetadata = null;
        this.stagedWorkbookReader?.DisposeWhenIdle();
        this.stagedWorkbookReader = null;
        this.loadingSheets.Clear();
        this.stagedSheetIndexes.Clear();
        this.previewedSheetIndexes.Clear();
        this.activeSheetLoadIndexes.Clear();
        this.failedSheetLoadIndexes.Clear();
        this.sheetDependencies.Clear();
        this.pendingUserRecalculations.Clear();
        this.deferredEditRecalculations.Clear();
        this.editDirtyNotifications.Clear();
        this.formulaBarEditSheet = null;
        this.pendingSheetActivationIndex = null;
        this.pendingWorksheetLink = null;
        this.CancelScheduledEditRecalculations();
    }

    private void StopStagedLoading(int generation, string placeholderText)
    {
        if (generation != this.workbookLoadGeneration)
        {
            return;
        }

        foreach (SpreadsheetViewModel sheet in this.loadingSheets)
        {
            CellViewModel? cell = sheet.GetCell(new CellAddress(0, 0));
            if (cell?.Text == Get(S.WorksheetLoadPrompt))
            {
                cell.Text = placeholderText;
            }
        }
        foreach (int sheetIndex in this.stagedSheetIndexes.Values)
        {
            this.failedSheetLoadIndexes.Add(sheetIndex);
        }
        this.ApplyReadOnly(raiseEvent: false);
        this.RebuildTabs();
    }

    // ----- Find and replace -----

    private void OnEditorKeyDown(object? sender, KeyEventArgs e)
    {
        if (!e.KeyModifiers.HasFlag(KeyModifiers.Control))
        {
            if (e.Key == Key.Escape && this.FindReplaceBar.IsVisible)
            {
                this.CloseFindReplace();
                e.Handled = true;
            }
            return;
        }

        if (e.Key == Key.F)
        {
            this.ShowFindReplace(focusReplace: false);
            e.Handled = true;
        }
        else if (e.Key == Key.H)
        {
            this.ShowFindReplace(focusReplace: true);
            e.Handled = true;
        }
    }

    private void OnFindBoxKeyDown(KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            this.FindMatch(e.KeyModifiers.HasFlag(KeyModifiers.Shift));
            e.Handled = true;
        }
        else if (e.Key == Key.Escape)
        {
            this.CloseFindReplace();
            e.Handled = true;
        }
    }

    private void ShowFindReplace(bool focusReplace)
    {
        this.FindReplaceBar.IsVisible = true;
        this.FindStatusText.Text = string.Empty;
        Dispatcher.UIThread.Post(() =>
        {
            TextBox target = focusReplace ? this.ReplaceTextBox : this.FindTextBox;
            target.Focus();
            target.SelectAll();
        }, DispatcherPriority.Input);
    }

    private void CloseFindReplace()
    {
        this.FindReplaceBar.IsVisible = false;
        this.FindStatusText.Text = string.Empty;
        this.Grid.FocusForSelection();
    }

    private void FindMatch(bool backwards)
    {
        string query = this.FindTextBox.Text ?? string.Empty;
        if (string.IsNullOrEmpty(query) || this.vm.RowCount == 0 || this.vm.ColumnCount == 0)
        {
            this.FindStatusText.Text = Get(S.EnterTextToFind);
            return;
        }

        int columns = this.vm.ColumnCount;
        int total = this.vm.RowCount * columns;
        int start = this.vm.ActiveCell.Row * columns + this.vm.ActiveCell.Column;
        int direction = backwards ? -1 : 1;
        for (int step = 1; step <= total; step++)
        {
            int index = (start + direction * step) % total;
            if (index < 0)
            {
                index += total;
            }

            CellAddress address = new(index / columns, index % columns);
            CellViewModel? cell = this.vm.GetCell(address);
            if (cell is not null && CellContains(cell, query))
            {
                this.Grid.SelectCell(address);
                this.FindStatusText.Text = address.ToA1();
                return;
            }
        }

        this.FindStatusText.Text = Get(S.NoMatches);
    }

    private static bool CellContains(CellViewModel cell, string query) =>
        cell.Text.Contains(query, StringComparison.CurrentCultureIgnoreCase) ||
        (!string.Equals(cell.DisplayText, cell.Text, StringComparison.Ordinal) &&
         cell.DisplayText.Contains(query, StringComparison.CurrentCultureIgnoreCase));

    private void ReplaceCurrentMatch()
    {
        string query = this.FindTextBox.Text ?? string.Empty;
        if (this.IsReadOnly || string.IsNullOrEmpty(query))
        {
            this.FindStatusText.Text = Get(this.IsReadOnly ? S.ReadOnly : S.EnterTextToFind);
            return;
        }

        CellViewModel? cell = this.vm.GetCell(this.vm.ActiveCell);
        if (cell is null || !cell.Text.Contains(query, StringComparison.CurrentCultureIgnoreCase))
        {
            this.FindMatch(backwards: false);
            return;
        }

        this.vm.Checkpoint();
        cell.Text = cell.Text.Replace(query, this.ReplaceTextBox.Text ?? string.Empty, StringComparison.CurrentCultureIgnoreCase);
        this.FindMatch(backwards: false);
    }

    private void ReplaceAllMatches()
    {
        string query = this.FindTextBox.Text ?? string.Empty;
        if (this.IsReadOnly || string.IsNullOrEmpty(query))
        {
            this.FindStatusText.Text = Get(this.IsReadOnly ? S.ReadOnly : S.EnterTextToFind);
            return;
        }

        string replacement = this.ReplaceTextBox.Text ?? string.Empty;
        List<(CellViewModel Cell, string Text, int Count)> changes = new();
        int total = 0;
        foreach (RowViewModel row in this.vm.Rows)
        {
            foreach (CellViewModel cell in row.Cells)
            {
                int count = CountOccurrences(cell.Text, query);
                if (count == 0)
                {
                    continue;
                }

                changes.Add((cell, cell.Text.Replace(query, replacement, StringComparison.CurrentCultureIgnoreCase), count));
                total += count;
            }
        }

        if (changes.Count == 0)
        {
            this.FindStatusText.Text = Get(S.NoMatches);
            return;
        }

        this.vm.Checkpoint();
        this.vm.RunBatch(() =>
        {
            foreach ((CellViewModel cell, string text, _) in changes)
            {
                cell.Text = text;
            }
        });
        this.FindStatusText.Text = Format(S.ReplacedCount, total);
    }

    private static int CountOccurrences(string text, string query)
    {
        int count = 0;
        for (int index = 0; index <= text.Length - query.Length;)
        {
            int match = text.IndexOf(query, index, StringComparison.CurrentCultureIgnoreCase);
            if (match < 0)
            {
                break;
            }

            count++;
            index = match + query.Length;
        }

        return count;
    }

    // Adds a new blank sheet at the end and switches to it.
    public void AddSheet()
    {
        SpreadsheetViewModel sheetVm = new() { Name = this.NextSheetName() };
        sheetVm.Initialize(this.InitialRowCount, this.InitialColumnCount);
        this.ConfigureWorkbookFormulaResolution(sheetVm);
        this.sheets.Add(sheetVm);
        this.ActivateSheet(this.sheets.Count - 1);
        this.RebuildTabs();
        this.ContentChanged?.Invoke(this, EventArgs.Empty);
    }

    // Loads a SheetData snapshot into a view model (names set before cells so =Name resolves).
    private static void ApplySheetData(SpreadsheetViewModel vm, SheetData sheet)
    {
        vm.SetNames(sheet.Names);
        // XLSX files already contain the last values calculated by Excel. Keep those cached results
        // during import; recalculating here both delays the first paint and can replace valid values
        // before every referenced worksheet is available.
        bool hasUncachedFormula = sheet.Cells
            .SelectMany(static row => row)
            .Any(static cell => FormulaEngine.IsFormula(cell.Text) && cell.FormulaResult is null);
        vm.LoadCellData(sheet.Cells, recalculate: hasUncachedFormula);
        for (int i = 0; i < vm.Columns.Count && i < sheet.ColumnWidths.Count; i++)
        {
            if (sheet.ColumnWidths[i] > 0)
            {
                vm.Columns[i].Width = sheet.ColumnWidths[i];
            }
        }

        vm.SetMerges(sheet.Merges);
        vm.SetImages(sheet.Images);
        vm.SetDataBars(sheet.DataBars);
        vm.SetCellValueRules(sheet.CellValueRules);
        vm.SetAutoFilterRange(sheet.AutoFilterRange);
        vm.SetHiddenRows(sheet.HiddenRows);
        vm.SetCharts(sheet.Charts);
        vm.TabColor = sheet.TabColor;
    }

    // ----- Overflow menu: spreadsheet-specific actions + formulas section -----

    private MenuFlyout BuildOverflowMenu()
    {
        MenuFlyout menu = new();

        MenuItem find = new() { Header = Get(S.FindReplaceMenu) };
        find.Click += (_, _) => this.ShowFindReplace(focusReplace: false);
        menu.Items.Add(find);
        menu.Items.Add(new Separator());

        MenuItem chart = new() { Header = Get(S.InsertChartMenu) };
        chart.Click += async (_, _) => await this.ShowChartBuilderAsync();
        menu.Items.Add(chart);

        MenuItem print = new() { Header = Get(S.PrintMenu) };
        print.Click += (_, _) => this.PrintPdf();
        menu.Items.Add(print);

        MenuItem autoSum = new() { Header = Get(S.AutoSumMenu) };
        autoSum.Click += (_, _) => this.AutoSum();
        menu.Items.Add(autoSum);

        MenuItem functions = new() { Header = Get(S.Functions) };
        foreach (MenuItem category in this.BuildFunctionMenuItems())
        {
            functions.Items.Add(category);
        }

        menu.Items.Add(functions);
        menu.Items.Add(new Separator());

        MenuItem addRow = new() { Header = Get(S.AddRow) };
        addRow.Click += (_, _) => this.AddRow();
        menu.Items.Add(addRow);

        MenuItem addColumn = new() { Header = Get(S.AddColumn) };
        addColumn.Click += (_, _) => this.AddColumn();
        menu.Items.Add(addColumn);

        return menu;
    }

    // The function library as a list of category MenuItems (each with its functions), reused by the
    // overflow menu's "Functions" submenu.
    private List<MenuItem> BuildFunctionMenuItems()
    {
        List<MenuItem> categories = new();
        foreach ((EditorString category, (string Name, EditorString Signature)[] functions) in FunctionLibrary)
        {
            MenuItem group = new() { Header = Get(category) };
            foreach ((string name, EditorString signature) in functions)
            {
                MenuItem item = new() { Header = name };
                ToolTip.SetTip(item, Get(signature));
                string fn = name;
                item.Click += (_, _) => this.InsertFunction(fn);
                group.Items.Add(item);
            }

            categories.Add(group);
        }

        return categories;
    }

    private void AddRow()
    {
        if (this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint();
        this.vm.EnsureSize(this.vm.RowCount + 1, this.vm.ColumnCount);
        this.Grid.Rebuild();
    }

    private void AddColumn()
    {
        if (this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint();
        this.vm.EnsureSize(this.vm.RowCount, this.vm.ColumnCount + 1);
        this.Grid.Rebuild();
    }

    // ----- Printing -----

    // Renders the active sheet to a temporary HTML file and opens it in the default browser, where the
    // user prints (or saves as PDF) via the browser's own print dialog. Avalonia has no built-in print
    // API, so this HTML-preview route is the portable, dependency-free way to print. Best-effort.
    public void Print()
    {
        try
        {
            string directory = Path.Combine(Path.GetTempPath(), "SpreadsheetEditor");
            Directory.CreateDirectory(directory);
            string file = Path.Combine(directory, SanitizeFileName(this.vm.Name) + ".html");
            File.WriteAllText(file, this.ExportHtml());
            Process.Start(new ProcessStartInfo { FileName = file, UseShellExecute = true });
        }
        catch
        {
            // Never crash the editor if the temp file or default browser can't be opened.
        }
    }

    // Writes the active sheet's used range to a paginated PDF (via SkiaSharp). Fully cross-platform
    // and self-contained — no browser or OS print subsystem required. Cell fills, borders, alignment,
    // and rich run formatting are preserved; the sheet is scaled to fit the page width and rows flow
    // across pages. Charts and images are not rendered (data only), matching ExportHtml.
    public void ExportPdf(Stream stream) => SpreadsheetPdfExporter.Export(this.vm, stream, this.Grid.RenderChartImages());

    // Convenience: renders the active sheet to a .pdf file at the given path.
    public void SavePdf(string path)
    {
        using FileStream stream = File.Create(path);
        this.ExportPdf(stream);
    }

    // ----- Zoom -----

    // The grid zoom level (1.0 = 100%). Scales cells, headers, and fonts; the underlying data and saved
    // column widths/row heights are unaffected.
    public double Zoom
    {
        get => this.Grid.Zoom;
        set => this.Grid.Zoom = value;
    }

    public void ZoomIn() => this.Grid.ZoomIn();

    public void ZoomOut() => this.Grid.ZoomOut();

    public void ResetZoom() => this.Grid.ResetZoom();

    // Message shown when the CUPS `lpr` command isn't available on Linux/macOS.
    private static string LprMissingMessage => Get(S.LprMissing);

    // Prints the active sheet. On Windows this shows the native Windows print dialog (comdlg32
    // PrintDlg) and prints the rendered pages to the chosen printer via GDI. On Linux/macOS it writes
    // a temporary PDF and submits it to the default CUPS printer via `lpr`, surfacing a message if
    // `lpr` is missing or fails.
    public void PrintPdf()
    {
        try
        {
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                this.PrintWindows();
                return;
            }

            string directory = Path.Combine(Path.GetTempPath(), "SpreadsheetEditor");
            Directory.CreateDirectory(directory);
            string file = Path.Combine(directory, SanitizeFileName(this.vm.Name) + ".pdf");
            this.SavePdf(file);
            _ = this.PrintUnixAsync(file);
        }
        catch (Exception ex)
        {
            _ = this.ShowMessageAsync(Get(S.PrintFailed), ex.Message);
        }
    }

    // Linux/macOS: submit the PDF to the default printer via CUPS' `lpr`. Reports a helpful message if
    // `lpr` isn't installed, and surfaces its error output (e.g. no default printer) when it fails.
    private async Task PrintUnixAsync(string file)
    {
        try
        {
            ProcessStartInfo info = new()
            {
                FileName = "lpr",
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true,
            };
            info.ArgumentList.Add(file);

            Process? process = Process.Start(info);
            if (process is null)
            {
                await this.ShowMessageAsync(Get(S.PrintFailed), LprMissingMessage);
                return;
            }

            string error = await process.StandardError.ReadToEndAsync();
            await process.WaitForExitAsync();
            if (process.ExitCode != 0)
            {
                await this.ShowMessageAsync(
                    Get(S.PrintFailed),
                    string.IsNullOrWhiteSpace(error)
                        ? Format(S.LprExitFailed, process.ExitCode)
                        : error.Trim());
            }
        }
        catch (Win32Exception)
        {
            // lpr isn't on PATH — the most common cause on a minimal Linux install.
            await this.ShowMessageAsync(Get(S.PrintFailed), LprMissingMessage);
        }
        catch (Exception ex)
        {
            await this.ShowMessageAsync(Get(S.PrintFailed), ex.Message);
        }
    }

    // A small modal message dialog (used for print errors). No-ops if there's no owner window.
    private async Task ShowMessageAsync(string title, string message)
    {
        if (TopLevel.GetTopLevel(this) is not Window owner)
        {
            return;
        }

        TextBlock text = new() { Text = message, TextWrapping = TextWrapping.Wrap, Margin = new Thickness(0, 0, 0, 12) };
        Button ok = new() { Content = Get(S.OK), IsDefault = true, IsCancel = true, HorizontalAlignment = HorizontalAlignment.Right };
        StackPanel panel = new() { Margin = new Thickness(16), Spacing = 8, Children = { text, ok } };
        Window dialog = new()
        {
            Title = title,
            Width = 420,
            SizeToContent = SizeToContent.Height,
            CanResize = false,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Content = panel,
        };
        ok.Click += (_, _) => dialog.Close();
        await dialog.ShowDialog(owner);
    }

    [SupportedOSPlatform("windows")]
    private void PrintWindows()
    {
        IntPtr owner = TopLevel.GetTopLevel(this)?.TryGetPlatformHandle()?.Handle ?? IntPtr.Zero;
        List<SkiaSharp.SKBitmap> pages = SpreadsheetPdfExporter.RenderPages(this.vm, this.Grid.RenderChartImages(), 200f);
        try
        {
            WindowsPrinter.PrintPages(owner, this.vm.Name, pages);
        }
        finally
        {
            foreach (SkiaSharp.SKBitmap page in pages)
            {
                page.Dispose();
            }
        }
    }

    // Builds a printable HTML rendering of the active sheet's used range, with per-cell formatting
    // (bold/italic/underline/strikethrough, colors, alignment, font) and merged regions preserved.
    public string ExportHtml()
    {
        SpreadsheetViewModel sheet = this.vm;
        int lastRow = Math.Max(sheet.LastNonEmptyRow(), 0);
        int lastColumn = LastNonEmptyColumn(sheet, lastRow);

        // Merge maps: anchor cell -> its span, and the set of covered cells to skip.
        Dictionary<(int Row, int Column), MergeRange> anchors = new();
        HashSet<(int Row, int Column)> covered = new();
        foreach (MergeRange merge in sheet.Merges)
        {
            anchors[(merge.Row, merge.Column)] = merge;
            for (int r = merge.Row; r < merge.Row + merge.RowSpan; r++)
            {
                for (int c = merge.Column; c < merge.Column + merge.ColumnSpan; c++)
                {
                    if (r != merge.Row || c != merge.Column)
                    {
                        covered.Add((r, c));
                    }
                }
            }
        }

        StringBuilder html = new();
        string title = WebUtility.HtmlEncode(sheet.Name);
        html.Append("<!doctype html><html><head><meta charset=\"utf-8\">");
        // Defense in depth: a strict CSP so no script executes and nothing is fetched, even if another
        // unencoded value is ever introduced. The exported document only needs inline styles + data: images.
        html.Append("<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; style-src 'unsafe-inline'; img-src data:\">");
        html.Append("<title>").Append(title).Append("</title><style>");
        html.Append("body{font-family:Calibri,'Segoe UI',sans-serif;margin:16px;}");
        html.Append("h1{font-size:15pt;margin:0 0 10px;}");
        html.Append("table{border-collapse:collapse;table-layout:fixed;}");
        html.Append("td{border:1px solid #c8c8c8;padding:2px 6px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}");
        html.Append("</style></head><body><h1>").Append(title).Append("</h1><table><colgroup>");
        for (int c = 0; c <= lastColumn; c++)
        {
            double width = c < sheet.ColumnCount ? sheet.Columns[c].Width : sheet.DefaultColumnWidth;
            html.Append("<col style=\"width:").Append(((int)width).ToString(CultureInfo.InvariantCulture)).Append("px\">");
        }

        html.Append("</colgroup>");
        for (int r = 0; r <= lastRow; r++)
        {
            html.Append("<tr>");
            for (int c = 0; c <= lastColumn; c++)
            {
                if (covered.Contains((r, c)))
                {
                    continue;
                }

                CellViewModel? cell = sheet.GetCell(new CellAddress(r, c));
                html.Append("<td");
                if (anchors.TryGetValue((r, c), out MergeRange span))
                {
                    if (span.ColumnSpan > 1)
                    {
                        html.Append(" colspan=\"").Append(span.ColumnSpan).Append('"');
                    }

                    if (span.RowSpan > 1)
                    {
                        html.Append(" rowspan=\"").Append(span.RowSpan).Append('"');
                    }
                }

                string style = CellStyle(cell);
                if (style.Length > 0)
                {
                    html.Append(" style=\"").Append(style).Append('"');
                }

                html.Append('>').Append(WebUtility.HtmlEncode(cell?.DisplayText ?? string.Empty)).Append("</td>");
            }

            html.Append("</tr>");
        }

        html.Append("</table></body></html>");
        return html.ToString();
    }

    // The last column (0-based) holding any content within the used rows, so blank trailing columns
    // aren't printed.
    private static int LastNonEmptyColumn(SpreadsheetViewModel sheet, int lastRow)
    {
        int last = 0;
        for (int r = 0; r <= lastRow && r < sheet.RowCount; r++)
        {
            RowViewModel row = sheet.Rows[r];
            for (int c = row.Cells.Count - 1; c > last; c--)
            {
                if (!string.IsNullOrEmpty(row.Cells[c].DisplayText))
                {
                    last = c;
                    break;
                }
            }
        }

        return last;
    }

    // Inline CSS for a cell's whole-cell formatting.
    private static string CellStyle(CellViewModel? cell)
    {
        if (cell is null)
        {
            return string.Empty;
        }

        StringBuilder style = new();
        if (cell.Bold)
        {
            style.Append("font-weight:bold;");
        }

        if (cell.Italic)
        {
            style.Append("font-style:italic;");
        }

        if (cell.Underline || cell.Strikethrough)
        {
            style.Append("text-decoration:");
            if (cell.Underline)
            {
                style.Append("underline ");
            }

            if (cell.Strikethrough)
            {
                style.Append("line-through");
            }

            style.Append(';');
        }

        if (cell.ForegroundColor is { } fg)
        {
            style.Append("color:").Append(ToHex(fg)).Append(';');
        }

        if (cell.BackgroundColor is { } bg)
        {
            style.Append("background-color:").Append(ToHex(bg)).Append(';');
        }

        string align = cell.TextAlignment switch
        {
            TextAlignment.Center => "center",
            TextAlignment.Right => "right",
            _ => "left",
        };
        style.Append("text-align:").Append(align).Append(';');
        style.Append("font-size:").Append(((int)Math.Round(cell.FontSize * 3.0 / 4.0)).ToString(CultureInfo.InvariantCulture)).Append("pt;");
        if (!string.IsNullOrWhiteSpace(cell.FontFamily))
        {
            // The font name is attacker-controlled (read verbatim from an imported .xlsx), so it must
            // not be emitted raw into a quoted CSS attribute — a quote would break out and inject markup.
            // Allowlist to characters valid in a CSS font-family name, then HTML-encode for the attribute
            // context (defense in depth: neither a quote nor a bracket survives both steps).
            string safeFamily = new string(cell.FontFamily
                .Where(static ch => char.IsLetterOrDigit(ch) || ch is ' ' or '-' or '_')
                .ToArray())
                .Trim();
            if (safeFamily.Length > 0)
            {
                style.Append("font-family:'").Append(WebUtility.HtmlEncode(safeFamily)).Append("';");
            }
        }

        return style.ToString();
    }

    private static string ToHex(Color color) => $"#{color.R:X2}{color.G:X2}{color.B:X2}";

    private static string SanitizeFileName(string name)
    {
        foreach (char invalid in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(invalid, '_');
        }

        return string.IsNullOrWhiteSpace(name) ? "sheet" : name;
    }

    // ----- Undo / redo (public, mirroring RichTextEditor.Undo/Redo) -----

    public void Undo()
    {
        if (!this.vm.CanUndo)
        {
            return;
        }

        CellAddress active = this.vm.ActiveCell;
        this.DetachActiveCell();
        this.vm.Undo();
        this.Grid.Rebuild();
        this.SetActiveCell(this.vm.IsValid(active) ? active : new CellAddress(0, 0));
        this.RefreshUndoButtons();
    }

    public void Redo()
    {
        if (!this.vm.CanRedo)
        {
            return;
        }

        CellAddress active = this.vm.ActiveCell;
        this.DetachActiveCell();
        this.vm.Redo();
        this.Grid.Rebuild();
        this.SetActiveCell(this.vm.IsValid(active) ? active : new CellAddress(0, 0));
        this.RefreshUndoButtons();
    }

    // Selects every cell in the active sheet (mirrors RichTextEditor.SelectAll).
    public void SelectAll() => this.Grid.SelectAll();

    // Subscribes to the active sheet's undo-state + content changes, unhooking the previous sheet so
    // events reflect only the active one.
    private void HookSheet()
    {
        if (this.hookedSheet is not null)
        {
            this.hookedSheet.UndoStateChanged -= this.RefreshUndoButtons;
            this.hookedSheet.DataChanged -= this.OnSheetContentChanged;
            this.hookedSheet.ViewChanged -= this.OnSheetViewChanged;
        }

        this.hookedSheet = this.vm;
        this.hookedSheet.UndoStateChanged += this.RefreshUndoButtons;
        this.hookedSheet.DataChanged += this.OnSheetContentChanged;
        this.hookedSheet.ViewChanged += this.OnSheetViewChanged;
        this.RefreshUndoButtons();
    }

    private void OnSheetViewChanged() => this.ContentChanged?.Invoke(this, EventArgs.Empty);

    private void OnSheetContentChanged()
    {
        if (!this.recalculatingWorkbookDependencies && this.hookedSheet is { } changedSheet)
        {
            this.RecalculateWorkbookFrom(changedSheet.Name, recalculateSource: false);
        }

        this.ContentChanged?.Invoke(this, EventArgs.Empty);
    }

    private void RefreshUndoButtons()
    {
        this.UndoButton.IsEnabled = this.vm.CanUndo;
        this.RedoButton.IsEnabled = this.vm.CanRedo;
    }

    // Inserts =NAME() into the active cell and puts the caret inside the parentheses
    // (in the formula bar) so the user can type the arguments.
    private void InsertFunction(string name)
    {
        if (this.activeCell is null || this.formulaBar is null || this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint();
        this.activeCell.Text = $"={name}()";
        this.Grid.SuspendEditor();
        this.formulaBar.Focus();
        this.formulaBar.CaretIndex = name.Length + 2; // just after "=NAME("
    }

    // Sums the contiguous numbers above the active cell (or to its left), into the active cell.
    private void AutoSum()
    {
        if (this.activeCell is null || this.formulaBar is null || this.IsReadOnly)
        {
            return;
        }

        // If a multi-cell range is selected, sum it; otherwise sum the run above/left of the cell.
        CellRange selection = this.Grid.SelectionRange;
        string? formula = selection.RowSpan > 1 || selection.ColumnSpan > 1
            ? $"=SUM({selection.ToA1()})"
            : this.SumRun(vertical: true) ?? this.SumRun(vertical: false);
        if (formula is null)
        {
            return;
        }

        this.vm.Checkpoint();
        this.activeCell.Text = formula;
        this.Grid.SuspendEditor();
        this.formulaBar.Focus();
        this.formulaBar.CaretIndex = formula.Length;
    }

    private string? SumRun(bool vertical)
    {
        CellAddress a = this.vm.ActiveCell;
        int i = (vertical ? a.Row : a.Column) - 1;
        while (i >= 0 && this.IsNumericCell(vertical ? new CellAddress(i, a.Column) : new CellAddress(a.Row, i)))
        {
            i--;
        }

        int start = i + 1;
        int end = (vertical ? a.Row : a.Column) - 1;
        if (start > end)
        {
            return null;
        }

        CellAddress from = vertical ? new CellAddress(start, a.Column) : new CellAddress(a.Row, start);
        CellAddress to = vertical ? new CellAddress(end, a.Column) : new CellAddress(a.Row, end);
        return $"=SUM({from.ToA1()}:{to.ToA1()})";
    }

    private bool IsNumericCell(CellAddress address)
    {
        CellViewModel? cell = this.vm.GetCell(address);
        return cell is not null && double.TryParse(cell.DisplayText,
            System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _);
    }

    // Opens the chart builder and inserts the chart the user configures.
    private async Task ShowChartBuilderAsync()
    {
        if (TopLevel.GetTopLevel(this) is not Window owner)
        {
            return;
        }

        // Default to the user's selected range if it spans multiple cells, else the current region.
        CellRange selection = this.Grid.SelectionRange;
        CellRange defaultRange;
        if (selection.RowSpan > 1 || selection.ColumnSpan > 1)
        {
            defaultRange = selection;
        }
        else
        {
            MergeRange region = this.vm.GetCurrentRegion(this.vm.ActiveCell);
            defaultRange = new CellRange(region.Row, region.Column, region.RowSpan, region.ColumnSpan);
        }

        ChartBuilderResult? result = await ChartBuilderDialog.ShowAsync(owner, defaultRange.ToA1());
        if (result is null)
        {
            return;
        }

        if (CellRange.TryParse(result.Range, out CellRange range))
        {
            this.vm.InsertChart(result.Kind, range, result.FirstRowHeaders, result.FirstColCategories,
                result.Title, result.ShowLegend, result.ShowDataLabels,
                result.ValueAxisFormat, result.AxisMinimum, result.AxisMaximum, result.AxisMajorStep);
        }
    }

    // Opens the builder pre-filled with a chart's settings and applies the user's changes.
    private async Task EditChartAsync(SheetChart chart)
    {
        if (TopLevel.GetTopLevel(this) is not Window owner)
        {
            return;
        }

        // Charts built from a contiguous block carry a SourceRange the dialog can re-read. Imported
        // charts define their series as explicit ranges (no single SourceRange) — for those, show the
        // first series' range and only apply setting changes, so editing never rebuilds (and wipes) the data.
        CellRange displayRange = chart.SourceRange
            ?? (chart.Series.Count > 0 ? chart.Series[0].PrimaryRange : new CellRange(chart.AnchorRow, 0, 1, 1));
        ChartBuilderResult initial = new(chart.Kind, displayRange.ToA1(), chart.Title,
            chart.FirstRowHeaders, chart.FirstColCategories, chart.ShowLegend, chart.ShowDataLabels,
            chart.ValueAxisFormat, chart.AxisMinimum, chart.AxisMaximum, chart.AxisMajorStep);
        // A chart built from a contiguous block keeps the block data-range + first-row/first-col options.
        // An imported chart's series are explicit ranges — expose each series' value range and the
        // category range as editable fields so its data source can be retargeted without a single block.
        bool blockBuilt = chart.SourceRange is not null;
        List<(string Title, string Range)>? seriesSources = blockBuilt ? null : DescribeChartSources(chart);

        // Pre-fill the category field with the range (if it comes from cells) or the literal labels
        // (Income, Expenses) so the user can see/edit them; blank only when the chart has no categories.
        string? categorySource = null;
        if (!blockBuilt)
        {
            if (chart.Categories is { } cats && HasRange(cats))
            {
                categorySource = cats.ToA1();
            }
            else if (chart.CachedCategories is { Count: > 0 } cached)
            {
                categorySource = string.Join(", ", cached);
            }
        }

        ChartBuilderResult? result = await ChartBuilderDialog.ShowAsync(
            owner, displayRange.ToA1(), initial, blockBuilt, seriesSources, categorySource);
        if (result is null)
        {
            return;
        }

        if (blockBuilt && CellRange.TryParse(result.Range, out CellRange newRange))
        {
            // Chart built from a block: rebuild from the (possibly edited) range.
            this.vm.UpdateChart(chart, result.Kind, newRange, result.FirstRowHeaders, result.FirstColCategories,
                result.Title, result.ShowLegend, result.ShowDataLabels,
                result.ValueAxisFormat, result.AxisMinimum, result.AxisMaximum, result.AxisMajorStep);
        }
        else
        {
            // Imported chart: apply the (possibly edited) series/category ranges plus the settings.
            this.vm.UpdateChartSource(chart, result.SeriesRanges, result.CategoryRange, result.Kind,
                result.Title, result.ShowLegend, result.ShowDataLabels,
                result.ValueAxisFormat, result.AxisMinimum, result.AxisMaximum, result.AxisMajorStep);
        }
    }

    private static bool HasRange(CellRange range) => range.RowSpan > 0 && range.ColumnSpan > 0;

    // Each series' title + its value range in A1 (empty when the series has no resolvable range), so the
    // edit dialog can show and let the user retarget where the chart's values come from.
    private static List<(string Title, string Range)> DescribeChartSources(SheetChart chart)
    {
        List<(string Title, string Range)> sources = new();
        foreach (ChartSeriesSpec series in chart.Series)
        {
            string name = string.IsNullOrWhiteSpace(series.Title) ? Get(S.Series) : series.Title;
            sources.Add((name, series.RangesA1));
        }

        return sources;
    }

    // The formula bar: a plain TextBox bound two-way to the active cell's text (its
    // DataContext is the active cell), so it shows and edits the cell content live.
    private void CreateFormulaBar()
    {
        this.formulaBar = new TextBox { VerticalContentAlignment = VerticalAlignment.Center };
        this.formulaBar.Bind(TextBox.TextProperty, new Binding(nameof(CellViewModel.Text))
        {
            Mode = BindingMode.TwoWay,
            UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged,
        });

        // Clicking into the formula bar moves the single cursor here: hide the in-cell editor so the
        // caret only exists in one place, and record a pre-edit undo checkpoint.
        this.formulaBar.GotFocus += (_, _) =>
        {
            this.vm.Checkpoint();
            this.Grid.SuspendEditor();
            this.formulaBarEditSheet = this.vm;
        };
        this.formulaBar.LostFocus += (_, _) =>
        {
            if (this.formulaBarEditSheet is { } editedSheet)
            {
                this.formulaBarEditSheet = null;
                this.CompleteDeferredEditRecalculation(editedSheet);
            }
        };
        this.EditorHost.Child = this.formulaBar;
    }

    // ----- Sheet tabs -----

    // Switches the editor to the sheet at index.
    private void ActivateSheet(int index)
    {
        if (index < 0 || index >= this.sheets.Count)
        {
            return;
        }

        if (this.formulaBarEditSheet is { } formulaSheet)
        {
            this.formulaBarEditSheet = null;
            this.CompleteDeferredEditRecalculation(formulaSheet);
        }
        this.Grid.SuspendEditor();
        this.DetachActiveCell();
        this.vm = this.sheets[index];
        this.DataContext = this.vm;
        this.Grid.Sheet = this.vm;            // rebinds + rebuilds the grid for this sheet
        this.SetActiveCell(new CellAddress(0, 0));
        this.HookSheet();
        this.ApplyReadOnly(raiseEvent: false);
    }

    private Task SwitchToSheetAsync(SpreadsheetViewModel sheet)
    {
        bool loading = this.loadingSheets.Contains(sheet);
        if (loading &&
            this.stagedSheetIndexes.TryGetValue(sheet, out int stagedIndex) &&
            !this.previewedSheetIndexes.Contains(stagedIndex) &&
            this.workbookLoadCancellation is { } previewCancellation)
        {
            // Keep the current frame on screen until the selected worksheet's viewport preview is
            // ready. The tab is highlighted immediately, then ReplaceStagedSheet performs one atomic
            // grid swap instead of briefly drawing the placeholder first.
            this.pendingSheetActivationIndex = stagedIndex;
            this.RebuildTabs();
            this.StartStagedSheetLoad(sheet, this.workbookLoadGeneration, previewCancellation.Token);
            return Task.CompletedTask;
        }

        this.pendingSheetActivationIndex = null;
        if (!ReferenceEquals(sheet, this.vm))
        {
            this.ActivateSheet(this.sheets.IndexOf(sheet));
            this.RebuildTabs();
        }

        if (loading && this.workbookLoadCancellation is { } loadCancellation)
        {
            this.StartStagedSheetLoad(sheet, this.workbookLoadGeneration, loadCancellation.Token);
        }

        return Task.CompletedTask;
    }

    private void NavigateWorksheetLink(string sheetName, CellAddress address)
    {
        SpreadsheetViewModel? target = this.sheets.FirstOrDefault(sheet =>
            string.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase));
        if (target is null)
        {
            return;
        }

        this.pendingWorksheetLink = (sheetName, address);
        _ = this.SwitchToSheetAsync(target);
        this.CompletePendingWorksheetLink(completed: false);
    }

    private void CompletePendingWorksheetLink(bool completed)
    {
        if (this.pendingWorksheetLink is not { } link ||
            !string.Equals(this.vm.Name, link.SheetName, StringComparison.OrdinalIgnoreCase))
        {
            return;
        }
        if (this.vm.IsValid(link.Address))
        {
            this.Grid.SelectCell(link.Address);
            this.pendingWorksheetLink = null;
        }
        else if (completed)
        {
            this.pendingWorksheetLink = null;
        }
    }

    private void RenameSheet(SpreadsheetViewModel sheet, string name)
    {
        name = name.Trim();
        if (string.IsNullOrEmpty(name) || name.Equals(sheet.Name, StringComparison.Ordinal))
        {
            return;
        }

        sheet.Name = this.UniqueName(name, except: sheet);
        this.RebuildTabs();
        this.ContentChanged?.Invoke(this, EventArgs.Empty);
    }

    private void DeleteSheet(SpreadsheetViewModel sheet)
    {
        if (this.sheets.Count <= 1)
        {
            return; // keep at least one sheet
        }

        int index = this.sheets.IndexOf(sheet);
        bool wasActive = ReferenceEquals(sheet, this.vm);
        this.sheets.Remove(sheet);
        this.loadingSheets.Remove(sheet);
        this.sheetDependencies.Remove(sheet);
        this.pendingUserRecalculations.Remove(sheet);
        this.deferredEditRecalculations.Remove(sheet);
        this.editDirtyNotifications.Remove(sheet);
        if (ReferenceEquals(this.formulaBarEditSheet, sheet))
        {
            this.formulaBarEditSheet = null;
        }
        this.CancelScheduledEditRecalculation(sheet);
        if (this.stagedSheetIndexes.Remove(sheet, out int stagedIndex))
        {
            if (this.pendingSheetActivationIndex == stagedIndex)
            {
                this.pendingSheetActivationIndex = null;
            }
            this.previewedSheetIndexes.Remove(stagedIndex);
            this.failedSheetLoadIndexes.Remove(stagedIndex);
        }
        if (wasActive)
        {
            this.ActivateSheet(Math.Min(index, this.sheets.Count - 1));
        }

        this.RebuildTabs();
        this.ContentChanged?.Invoke(this, EventArgs.Empty);
    }

    private string NextSheetName()
    {
        string prefix = Get(S.Sheet);
        long next = 1;
        foreach (SpreadsheetViewModel sheet in this.sheets)
        {
            if (sheet.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
                long.TryParse(sheet.Name.AsSpan(prefix.Length), NumberStyles.None,
                    CultureInfo.InvariantCulture, out long number) && number >= next)
            {
                next = number + 1;
            }
        }
        return $"{prefix}{next}";
    }

    // Returns desired, suffixed (e.g. "Sheet2") to stay unique among sheets.
    private string UniqueName(string desired, SpreadsheetViewModel? except = null)
    {
        desired = string.IsNullOrWhiteSpace(desired) ? Get(S.Sheet) : desired.Trim();
        bool Taken(string n) => this.sheets.Any(s => !ReferenceEquals(s, except) && s.Name.Equals(n, StringComparison.OrdinalIgnoreCase));
        if (!Taken(desired))
        {
            return desired;
        }

        for (int i = 2; ; i++)
        {
            string candidate = $"{desired}{i}";
            if (!Taken(candidate))
            {
                return candidate;
            }
        }
    }

    // Rebuilds the bottom tab strip from the sheet list (active tab highlighted).
    private void RebuildTabs()
    {
        this.SheetTabs.Children.Clear();
        foreach (SpreadsheetViewModel sheet in this.sheets)
        {
            this.SheetTabs.Children.Add(this.BuildTab(sheet));
        }
    }

    private Control BuildTab(SpreadsheetViewModel sheet)
    {
        bool pending = this.loadingSheets.Contains(sheet);
        bool staged = this.stagedSheetIndexes.TryGetValue(sheet, out int sheetIndex);
        bool active = ReferenceEquals(sheet, this.vm) ||
            (staged && this.pendingSheetActivationIndex == sheetIndex);
        bool loading = staged && this.activeSheetLoadIndexes.Contains(sheetIndex);
        bool failed = staged && this.failedSheetLoadIndexes.Contains(sheetIndex);
        bool dark = this.ActualThemeVariant == ThemeVariant.Dark;
        TextBlock text = new()
        {
            Text = loading ? $"{sheet.Name} …" : sheet.Name,
            Opacity = 1,
            VerticalAlignment = VerticalAlignment.Center,
            FontWeight = active ? FontWeight.SemiBold : FontWeight.Normal,
        };
        Color? tabColor = sheet.TabColor;
        if (active)
        {
            // Do not rely on a host's Button foreground theme here: the active tab uses an explicit
            // surface and must keep its label readable in both light and dark mode.
            text.Foreground = dark ? Brushes.White : Brushes.Black;
        }
        else if (tabColor is { } color)
        {
            // Excel fills inactive colored tabs and switches the label for sufficient contrast.
            double luminance = 0.2126 * color.R + 0.7152 * color.G + 0.0722 * color.B;
            text.Foreground = luminance < 145 ? Brushes.White : Brushes.Black;
        }

        // Active tab reads as the "front" sheet: an elevated surface with its imported tab color as
        // the accent edge (or the editor's blue accent when the workbook has no tab color).
        IBrush activeBackground = dark ? new SolidColorBrush(Color.FromRgb(0x3C, 0x3C, 0x3C)) : Brushes.White;
        IBrush accent = tabColor is { } tc
            ? new SolidColorBrush(tc)
            : new SolidColorBrush(Color.FromRgb(0x1A, 0x73, 0xE8));
        Button tab = new()
        {
            Content = text,
            Focusable = false,
            Padding = new Thickness(10, 3),
            Background = active
                ? activeBackground
                : tabColor is { } inactiveColor ? new SolidColorBrush(inactiveColor) : Brushes.Transparent,
            BorderThickness = new Thickness(1, active ? 2 : 1, 1, 0),
            BorderBrush = active ? accent : Brushes.Transparent,
        };
        if (failed)
        {
            ToolTip.SetTip(tab, Get(S.WorksheetLoadFailedTip));
        }
        else if (loading)
        {
            ToolTip.SetTip(tab, Get(S.WorksheetLoadingTip));
        }
        else if (pending)
        {
            ToolTip.SetTip(tab, Get(S.WorksheetLoadTip));
        }
        tab.Click += async (_, _) => await this.SwitchToSheetAsync(sheet);

        // Double-click a tab to rename it inline.
        tab.AddHandler(PointerPressedEvent, (_, e) =>
        {
            if (e.ClickCount >= 2)
            {
                this.BeginRenameTab(sheet, tab);
            }
        }, Avalonia.Interactivity.RoutingStrategies.Tunnel);

        // Right-click: Rename / Delete.
        MenuFlyout menu = new();
        MenuItem rename = new() { Header = Get(S.Rename) };
        rename.Click += (_, _) => this.BeginRenameTab(sheet, tab);
        MenuItem delete = new() { Header = Get(S.Delete), IsEnabled = this.sheets.Count > 1 };
        delete.Click += (_, _) => this.DeleteSheet(sheet);
        menu.Items.Add(rename);
        menu.Items.Add(delete);
        tab.ContextFlyout = menu;
        return tab;
    }

    // Swaps a tab's label for an inline editor to rename the sheet.
    private void BeginRenameTab(SpreadsheetViewModel sheet, Button tab)
    {
        TextBox box = new()
        {
            Text = sheet.Name,
            MinWidth = 70,
            Padding = new Thickness(2, 0),
        };
        tab.Content = box;
        void Commit()
        {
            this.RenameSheet(sheet, box.Text ?? sheet.Name);
            this.RebuildTabs();
        }

        box.KeyDown += (_, e) =>
        {
            if (e.Key == Avalonia.Input.Key.Enter) { Commit(); e.Handled = true; }
            else if (e.Key == Avalonia.Input.Key.Escape) { this.RebuildTabs(); e.Handled = true; }
        };
        box.LostFocus += (_, _) => Commit();
        Dispatcher.UIThread.Post(() => { box.Focus(); box.SelectAll(); });
    }

    // ----- Active cell -----

    private void SetActiveCell(CellAddress address)
    {
        if (!this.vm.IsValid(address))
        {
            return;
        }

        if (this.activeCell is not null && this.vm.ActiveCell.Equals(address) &&
            ReferenceEquals(this.activeCell, this.vm.GetCell(address)))
        {
            return;
        }

        this.DetachActiveCell();

        this.vm.ActiveCell = address;
        this.activeCell = this.vm.GetCell(address);
        this.NameBox.Text = address.ToA1();

        if (this.activeCell is null)
        {
            return;
        }

        this.activeCell.IsActive = true;
        if (this.formulaBar is not null)
        {
            this.formulaBar.DataContext = this.activeCell;
        }

        this.SyncToolbar(this.activeCell);
        this.SelectionChanged?.Invoke(this, EventArgs.Empty);
    }

    private void DetachActiveCell()
    {
        if (this.formulaBar is not null)
        {
            this.formulaBar.DataContext = null;
        }

        if (this.activeCell is null)
        {
            return;
        }

        this.activeCell.IsActive = false;
        this.activeCell = null;
    }

    // ----- Toolbar formatting (applies to every cell in the selection) -----

    private void ApplyToSelection(Action<CellViewModel> apply)
    {
        if (this.suppressToolbar || this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint(); // whole-selection formatting is one undo step
        foreach (CellViewModel cell in this.Grid.SelectionCells())
        {
            apply(cell);
        }
    }

    private void ApplyForeground(Color color)
    {
        this.ForegroundSwatch.Background = new SolidColorBrush(color);
        if (this.suppressToolbar)
        {
            return;
        }

        // Route through the editor: per-character when editing, else whole-selection.
        this.Grid.ApplyForegroundToCell(color);
    }

    private void ApplyBackground(Color color)
    {
        this.BackgroundSwatch.Background = new SolidColorBrush(color);
        this.ApplyToSelection(c => c.BackgroundColor = color);
    }

    private void SyncToolbar(CellViewModel cell)
    {
        this.suppressToolbar = true;
        try
        {
            this.BoldButton.IsChecked = cell.Bold;
            this.ItalicButton.IsChecked = cell.Italic;
            this.UnderlineButton.IsChecked = cell.Underline;
            this.StrikethroughButton.IsChecked = cell.Strikethrough;
            this.ForegroundSwatch.Background = cell.Foreground ?? Brushes.Black;
            this.BackgroundSwatch.Background = cell.Background ?? Brushes.Transparent;

            this.FontFamilyBox.SelectedItem = cell.FontFamily ?? DefaultFontName;
            this.FontSizeBox.SelectedItem = (int)Math.Round(cell.FontSize * 3.0 / 4.0); // px → pt
            this.AlignLeftButton.IsChecked = cell.TextAlignment == TextAlignment.Left;
            this.AlignCenterButton.IsChecked = cell.TextAlignment == TextAlignment.Center;
            this.AlignRightButton.IsChecked = cell.TextAlignment == TextAlignment.Right;
        }
        finally
        {
            this.suppressToolbar = false;
        }
    }

    private void SetupFontControls()
    {
        this.FontFamilyBox.ItemsSource = FontNames;
        this.FontSizeBox.ItemsSource = FontSizes;
        this.FontFamilyBox.SelectionChanged += (_, _) =>
        {
            if (!this.suppressToolbar && this.FontFamilyBox.SelectedItem is string name)
            {
                this.ApplyToSelection(c => c.FontFamily = name);
            }
        };
        this.FontSizeBox.SelectionChanged += (_, _) =>
        {
            if (!this.suppressToolbar && this.FontSizeBox.SelectedItem is int pt)
            {
                this.ApplyToSelection(c => c.FontSize = pt * 4.0 / 3.0); // pt → px
            }
        };
    }

    // Reflects the current B/I/U state (editor selection while editing, else whole-cell) on the toolbar.
    private void SyncFormatButtons()
    {
        (bool bold, bool italic, bool underline, bool strikethrough) = this.Grid.FormatState();
        this.suppressToolbar = true;
        try
        {
            this.BoldButton.IsChecked = bold;
            this.ItalicButton.IsChecked = italic;
            this.UnderlineButton.IsChecked = underline;
            this.StrikethroughButton.IsChecked = strikethrough;
        }
        finally
        {
            this.suppressToolbar = false;
        }
    }

    private void SetAlignment(TextAlignment alignment)
    {
        this.ApplyToSelection(c => c.TextAlignment = alignment);
        this.suppressToolbar = true;
        this.AlignLeftButton.IsChecked = alignment == TextAlignment.Left;
        this.AlignCenterButton.IsChecked = alignment == TextAlignment.Center;
        this.AlignRightButton.IsChecked = alignment == TextAlignment.Right;
        this.suppressToolbar = false;
    }

    private void BuildPalette(WrapPanel panel, Action<Color> apply)
    {
        foreach (Color color in Palette)
        {
            Button swatch = new()
            {
                Width = 28,
                Height = 28,
                Margin = new Thickness(1),
                Padding = new Thickness(0),
                Background = new SolidColorBrush(color),
                BorderThickness = new Thickness(1),
                BorderBrush = Brushes.Gray,
                Focusable = false,
            };
            Color captured = color;
            swatch.Click += (_, _) => apply(captured);
            panel.Children.Add(swatch);
        }
    }

    // ----- Insert special characters (currency + math symbols) -----

    private static readonly string[] CurrencySymbols =
    {
        "$", "€", "£", "¥", "¢", "₹", "₩",
        "₽", "₪", "₺", "₫", "฿", "₦", "₱",
    };

    private static readonly string[] MathSymbols =
    {
        "+", "−", "×", "÷", "=", "≠", "≈", "±",
        "√", "∑", "∏", "∫", "∞", "≤", "≥",
        "π", "µ", "°", "‰", "²", "³",
        "½", "¼", "¾", "·", "Δ",
    };

    private Flyout BuildSymbolFlyout()
    {
        StackPanel root = new() { Spacing = 6, MaxWidth = 270 };
        root.Children.Add(this.SymbolSection(Get(S.Currency), CurrencySymbols));
        root.Children.Add(this.SymbolSection(Get(S.Math), MathSymbols));
        return new Flyout { Content = root };
    }

    private Control SymbolSection(string title, string[] symbols)
    {
        StackPanel section = new() { Spacing = 2 };
        section.Children.Add(new TextBlock { Text = title, FontWeight = FontWeight.SemiBold, Margin = new Thickness(2, 0) });
        WrapPanel grid = new() { MaxWidth = 264 };
        foreach (string symbol in symbols)
        {
            Button button = new()
            {
                Content = symbol,
                Width = 30,
                Height = 28,
                Margin = new Thickness(1),
                Padding = new Thickness(0),
                Focusable = false,
                FontSize = 15,
            };
            ToolTip.SetTip(button, symbol);
            string captured = symbol;
            button.Click += (_, _) => this.InsertSymbol(captured);
            grid.Children.Add(button);
        }

        section.Children.Add(grid);
        return section;
    }

    // Appends a symbol to the active cell's content (via the formula bar), as one undo step.
    private void InsertSymbol(string symbol)
    {
        if (this.activeCell is null || this.formulaBar is null || this.IsReadOnly)
        {
            return;
        }

        this.vm.Checkpoint();
        this.Grid.SuspendEditor();
        this.formulaBar.Focus();
        string text = this.formulaBar.Text ?? string.Empty;
        this.formulaBar.Text = text + symbol;
        this.formulaBar.CaretIndex = (this.formulaBar.Text ?? string.Empty).Length;
    }
}
