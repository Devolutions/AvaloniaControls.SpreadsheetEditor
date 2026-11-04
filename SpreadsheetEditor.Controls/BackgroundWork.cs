namespace SpreadsheetEditor.Controls;

using System;
using System.Threading;
using System.Threading.Tasks;

// CPU-heavy workbook work must not compete with input/rendering at normal thread priority. Use a
// dedicated worker: changing a pooled thread's priority could affect unrelated host application work.
internal static class BackgroundWork
{
    internal static Task<T> Run<T>(Func<T> work, CancellationToken cancellationToken)
    {
        if (cancellationToken.IsCancellationRequested) return Task.FromCanceled<T>(cancellationToken);
        // Completing a Task directly on the low-priority worker can inline its caller's await
        // continuation on that worker. Keep the worker's thread priority confined to CPU work.
        TaskCompletionSource<T> completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
        _ = Task.Factory.StartNew(() =>
        {
            try
            {
                Thread.CurrentThread.Name ??= "Spreadsheet calculation";
                Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
                cancellationToken.ThrowIfCancellationRequested();
                completion.TrySetResult(work());
            }
            catch (OperationCanceledException exception) { completion.TrySetCanceled(exception.CancellationToken); }
            catch (Exception exception) { completion.TrySetException(exception); }
        }, CancellationToken.None, TaskCreationOptions.LongRunning, TaskScheduler.Default);
        return completion.Task;
    }
}
