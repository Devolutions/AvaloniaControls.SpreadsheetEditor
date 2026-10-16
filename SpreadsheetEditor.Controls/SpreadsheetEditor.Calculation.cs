namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

using Avalonia.Threading;

using global::SpreadsheetEditor.Controls.Formulas;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.ViewModels;

public partial class SpreadsheetEditor
{
    private readonly HashSet<string> pendingCalculationSheets = new(StringComparer.OrdinalIgnoreCase);
    private CancellationTokenSource? workbookCalculationCancellation;

    private sealed record CalculationSheet(
        SpreadsheetViewModel Model,
        long Version,
        string Name,
        RowViewModel[] Rows,
        string[][] Values,
        Dictionary<CellAddress, string?> Formats,
        Dictionary<string, string> Names,
        CellAddress[] Formulas);

    private sealed record CalculationResult(CalculationSheet Sheet, List<(CellAddress Address, CellValue Value, string Display)> Cells);

    private static bool IsChartInput(SpreadsheetViewModel sheet, CellAddress address)
    {
        bool Contains(CellRange range) => address.Row >= range.Row && address.Row < range.Row + range.RowSpan &&
            address.Column >= range.Column && address.Column < range.Column + range.ColumnSpan;
        return sheet.Charts.Any(chart =>
            (chart.SourceRange is { } source && Contains(source)) ||
            (chart.Categories is { } categories && Contains(categories)) ||
            chart.Series.Any(series => series.Ranges.Any(Contains)));
    }

    private void CancelWorkbookCalculation(bool clearPending)
    {
        // The operation owns disposal after its worker has observed cancellation.
        this.workbookCalculationCancellation?.Cancel();
        this.workbookCalculationCancellation = null;
        if (clearPending)
        {
            this.pendingCalculationSheets.Clear();
        }
    }

    private void StartPendingWorkbookCalculation()
    {
        if (this.workbookCalculationCancellation is not null || this.pendingCalculationSheets.Count == 0 ||
            this.sheets.Any(this.IsActivelyEditing))
        {
            return;
        }

        CancellationTokenSource cancellation = new();
        this.workbookCalculationCancellation = cancellation;
        _ = this.CalculateWorkbookAsync(cancellation);
    }

    private async Task CalculateWorkbookAsync(CancellationTokenSource cancellation)
    {
        CancellationToken token = cancellation.Token;
        CalculationSheet[] snapshots = Array.Empty<CalculationSheet>();
        HashSet<string> targets = new(StringComparer.OrdinalIgnoreCase);
        try
        {
            // Always yield before capturing inputs, including when called by a load-completion event.
            await Task.Delay(1, token).ConfigureAwait(false);
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                token.ThrowIfCancellationRequested();
                targets.UnionWith(this.pendingCalculationSheets);
                snapshots = this.sheets.Where(sheet => !this.loadingSheets.Contains(sheet))
                    .Select(sheet => new CalculationSheet(
                        sheet, sheet.CalculationVersion, sheet.Name, sheet.Rows.ToArray(),
                        new string[sheet.RowCount][],
                        new Dictionary<CellAddress, string?>(),
                        new Dictionary<string, string>(sheet.Names, StringComparer.OrdinalIgnoreCase),
                        sheet.FormulaAddresses()))
                    .ToArray();
            }, DispatcherPriority.Background);

            foreach (CalculationSheet snapshot in snapshots)
            {
                int row = 0;
                int column = 0;
                while (row < snapshot.Rows.Length)
                {
                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        this.ValidateCalculation(snapshots, token);
                        Stopwatch budget = Stopwatch.StartNew();
                        do
                        {
                            RowViewModel input = snapshot.Rows[row];
                            snapshot.Values[row] ??= new string[input.Cells.Count];
                            while (column < input.Cells.Count)
                            {
                                CellViewModel cell = input.Cells[column];
                                snapshot.Values[row][column] = cell.Text;
                                if (cell.IsFormula)
                                {
                                    snapshot.Formats[new CellAddress(row, column)] = cell.NumberFormat;
                                }
                                column++;
                                // Wide sheets need a budget within a row, not just between rows.
                                if ((column & 63) == 0 && budget.ElapsedMilliseconds >= 4) return;
                            }
                            row++;
                            column = 0;
                        }
                        while (row < snapshot.Rows.Length && budget.ElapsedMilliseconds < 4);
                    }, DispatcherPriority.Background);
                    if (row < snapshot.Rows.Length)
                    {
                        await Task.Delay(1, token).ConfigureAwait(false);
                    }
                }
            }

            // Only immutable text/name arrays and private formula engines are used by this worker.
            // No observable model, control or live sheet engine is read or modified off the UI thread.
            List<CalculationResult> results = await BackgroundWork.Run(
                () => CalculateSnapshots(snapshots, targets, token), token).ConfigureAwait(false);
            foreach (CalculationResult result in results)
            {
                int index = 0;
                while (index < result.Cells.Count)
                {
                    await Dispatcher.UIThread.InvokeAsync(() =>
                    {
                        this.ValidateCalculation(snapshots, token);
                        this.Grid.ApplyCalculatedCells(() =>
                        {
                            Stopwatch budget = Stopwatch.StartNew();
                            do
                            {
                                (CellAddress address, CellValue value, string display) = result.Cells[index++];
                                result.Sheet.Model.ApplyCalculatedValue(address, value, display);
                            }
                            while (index < result.Cells.Count && budget.ElapsedMilliseconds < 4);
                        });
                    }, DispatcherPriority.Background);
                    if (index < result.Cells.Count)
                    {
                        await Task.Delay(1, token).ConfigureAwait(false);
                    }
                }
            }

            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                this.ValidateCalculation(snapshots, token);
                this.pendingCalculationSheets.ExceptWith(targets);
                this.recalculatingWorkbookDependencies = true;
                try
                {
                    foreach (CalculationResult result in results)
                    {
                        result.Sheet.Model.CompleteBackgroundCalculation();
                    }
                    if (results.Any(result => ReferenceEquals(result.Sheet.Model, this.vm)))
                    {
                        this.Grid.RefreshCalculatedCells();
                    }
                }
                finally
                {
                    this.recalculatingWorkbookDependencies = false;
                }
            }, DispatcherPriority.Background);
        }
        catch (OperationCanceledException)
        {
            // New edits, replacement sheets and new workbooks supersede their old snapshots.
        }
        catch (Exception exception)
        {
            // Preserve the last displayed values if an unsupported calculation fails unexpectedly.
            Trace.TraceError("Spreadsheet background calculation failed: {0}", exception);
        }
        finally
        {
            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                if (ReferenceEquals(this.workbookCalculationCancellation, cancellation))
                {
                    this.workbookCalculationCancellation = null;
                }
                cancellation.Dispose();
            }, DispatcherPriority.Background);
        }
    }

    private void ValidateCalculation(IReadOnlyList<CalculationSheet> snapshots, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (snapshots.Any(snapshot => snapshot.Version != snapshot.Model.CalculationVersion ||
            !snapshot.Name.Equals(snapshot.Model.Name, StringComparison.Ordinal) || !this.sheets.Contains(snapshot.Model)))
        {
            throw new OperationCanceledException(token);
        }
    }

    private static List<CalculationResult> CalculateSnapshots(
        IReadOnlyList<CalculationSheet> snapshots,
        IReadOnlySet<string> targets,
        CancellationToken token)
    {
        Dictionary<string, FormulaEngine> engines = new(StringComparer.OrdinalIgnoreCase);
        foreach (CalculationSheet snapshot in snapshots)
        {
            FormulaEngine engine = new(
                address => address.Row >= 0 && address.Row < snapshot.Values.Length &&
                    address.Column >= 0 && address.Column < snapshot.Values[address.Row].Length
                        ? snapshot.Values[address.Row][address.Column] : string.Empty,
                name => snapshot.Names.GetValueOrDefault(name),
                (sheetName, address) => engines.TryGetValue(sheetName ?? snapshot.Name, out FormulaEngine? referenced)
                    ? referenced.Evaluate(address) : CellValue.FromError(FormulaErrors.Ref))
            { CancellationToken = token };
            engines.Add(snapshot.Name, engine);
        }

        List<CalculationResult> results = new();
        foreach (CalculationSheet snapshot in snapshots.Where(snapshot => targets.Contains(snapshot.Name)))
        {
            List<(CellAddress, CellValue, string)> cells = new(snapshot.Formulas.Length);
            foreach (CellAddress address in snapshot.Formulas)
            {
                token.ThrowIfCancellationRequested();
                CellValue value = engines[snapshot.Name].Evaluate(address);
                string display = value.Kind == CellValueKind.Number
                    ? NumberFormatter.Format(value.Number, snapshot.Formats.GetValueOrDefault(address)) : value.Display();
                cells.Add((address, value, display));
            }
            results.Add(new CalculationResult(snapshot, cells));
        }
        return results;
    }
}
