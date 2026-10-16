namespace SpreadsheetEditor.Controls;

using System;
using System.Threading;
using System.Threading.Tasks;

// CPU-heavy workbook work must not compete with input/rendering at normal thread priority. Use a
// dedicated worker: changing a pooled thread's priority could affect unrelated host application work.
internal static class BackgroundWork
{
    internal static Task<T> Run<T>(Func<T> work, CancellationToken cancellationToken) =>
        Task.Factory.StartNew(() =>
        {
            Thread.CurrentThread.Name ??= "Spreadsheet calculation";
            Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
            cancellationToken.ThrowIfCancellationRequested();
            return work();
        }, cancellationToken, TaskCreationOptions.LongRunning, TaskScheduler.Default);
}
