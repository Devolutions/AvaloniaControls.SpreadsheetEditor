namespace SampleApp;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using SpreadsheetEditor.Controls.Localization;

public partial class MainWindow : Window
{
    private static readonly FilePickerFileType ExcelFiles = new(EditorText.Get(EditorString.ExcelFileFilter))
    {
        Patterns = new[] { "*.xlsx" },
        MimeTypes = new[] { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        AppleUniformTypeIdentifiers = new[] { "org.openxmlformats.spreadsheetml.sheet" },
    };

    private static readonly FilePickerFileType TextFiles = new(EditorText.Get(EditorString.TextFileFilter))
    {
        Patterns = new[] { "*.tsv", "*.csv", "*.txt" },
    };

    private static readonly FilePickerFileType PdfFiles = new(EditorText.Get(EditorString.PdfFileFilter))
    {
        Patterns = new[] { "*.pdf" },
        MimeTypes = new[] { "application/pdf" },
        AppleUniformTypeIdentifiers = new[] { "com.adobe.pdf" },
    };

    private static bool IsXlsx(string name) =>
        name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase);

    private IStorageFile? currentFile;

    public MainWindow()
    {
        this.InitializeComponent();

        this.NewItem.Click += this.OnNew;
        this.OpenItem.Click += this.OnOpen;
        this.SaveItem.Click += this.OnSave;
        this.SaveAsItem.Click += this.OnSaveAs;
        this.ExportPdfItem.Click += this.OnExportPdf;
        this.PrintItem.Click += this.OnPrint;
        this.CloseItem.Click += this.OnClose;

        // Ctrl+P prints (the grid's key handler ignores it, so it bubbles up to here).
        this.KeyDown += (_, e) =>
        {
            if (e.Key == Key.P && e.KeyModifiers == KeyModifiers.Control)
            {
                this.Editor.PrintPdf();
                e.Handled = true;
            }
        };

        this.UpdateTitle();
    }

    private void OnPrint(object? sender, RoutedEventArgs e) => this.Editor.PrintPdf();

    private async void OnExportPdf(object? sender, RoutedEventArgs e)
    {
        IStorageFile? file = await this.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
        {
            Title = EditorText.Get(EditorString.ExportPdfTitle),
            DefaultExtension = "pdf",
            SuggestedFileName = EditorText.Get(EditorString.SampleSheetPdfName),
            FileTypeChoices = new[] { PdfFiles },
        });

        if (file is null)
        {
            return;
        }

        await using Stream stream = await file.OpenWriteAsync();
        this.Editor.ExportPdf(stream);
    }

    private void OnNew(object? sender, RoutedEventArgs e)
    {
        this.Editor.NewSheet(50, 26);
        this.currentFile = null;
        this.UpdateTitle();
    }

    private void OnClose(object? sender, RoutedEventArgs e)
    {
        this.Editor.NewSheet(50, 26);
        this.currentFile = null;
        this.UpdateTitle();
    }

    private async void OnOpen(object? sender, RoutedEventArgs e)
    {
        try
        {
            IReadOnlyList<IStorageFile> files = await this.StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = EditorText.Get(EditorString.OpenSpreadsheetTitle),
                AllowMultiple = false,
                FileTypeFilter = new[] { ExcelFiles, TextFiles },
            });

            if (files.Count == 0)
            {
                return;
            }

            IStorageFile file = files[0];
            await using Stream stream = await file.OpenReadAsync();

            if (IsXlsx(file.Name))
            {
                await this.Editor.OpenXlsxAsync(stream);
            }
            else
            {
                using StreamReader reader = new(stream);
                string text = await reader.ReadToEndAsync();
                this.Editor.LoadRows(SpreadsheetFile.Parse(text, SpreadsheetFile.DelimiterFor(file.Name)));
            }

            this.currentFile = file;
            this.UpdateTitle();
        }
        catch (Exception ex)
        {
            await this.ShowError(EditorText.Get(EditorString.OpenFailed), ex.Message);
        }
    }

    private async void OnSave(object? sender, RoutedEventArgs e)
    {
        if (this.currentFile is null)
        {
            await this.SaveAs();
            return;
        }

        await this.WriteTo(this.currentFile);
    }

    private async void OnSaveAs(object? sender, RoutedEventArgs e) => await this.SaveAs();

    private async Task SaveAs()
    {
        try
        {
            IStorageFile? file = await this.StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions
            {
                Title = EditorText.Get(EditorString.SaveSpreadsheetTitle),
                DefaultExtension = "xlsx",
                SuggestedFileName = EditorText.Get(EditorString.SampleSheetXlsxName),
                FileTypeChoices = new[] { ExcelFiles, TextFiles },
            });

            if (file is null)
            {
                return;
            }

            await this.WriteTo(file);
            this.currentFile = file;
            this.UpdateTitle();
        }
        catch (Exception ex)
        {
            await this.ShowError(EditorText.Get(EditorString.SaveFailed), ex.Message);
        }
    }

    private async Task WriteTo(IStorageFile file)
    {
        // Build the full file content first.
        byte[] bytes;
        if (IsXlsx(file.Name))
        {
            using MemoryStream buffer = new();
            this.Editor.SaveXlsx(buffer);
            bytes = buffer.ToArray();
        }
        else
        {
            string text = SpreadsheetFile.ToDelimited(this.Editor.GetRows(), SpreadsheetFile.DelimiterFor(file.Name));
            bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(text);
        }

        // On desktop the picker returns a real path — write directly so the file is fully replaced
        // (OpenWriteAsync doesn't truncate, which corrupts a shorter .xlsx/zip on overwrite).
        string? localPath = file.TryGetLocalPath();
        if (localPath is not null)
        {
            await File.WriteAllBytesAsync(localPath, bytes);
            return;
        }

        await using Stream stream = await file.OpenWriteAsync();
        await stream.WriteAsync(bytes);
        if (stream.CanSeek)
        {
            stream.SetLength(bytes.Length);
        }
    }

    private void UpdateTitle() =>
        this.Title = this.currentFile is null
            ? EditorText.Get(EditorString.SampleTitle)
            : $"{EditorText.Get(EditorString.SampleTitle)} — {this.currentFile.Name}";

    private async Task ShowError(string title, string message)
    {
        Window dialog = new()
        {
            Title = title,
            Width = 380,
            Height = 140,
            WindowStartupLocation = WindowStartupLocation.CenterOwner,
            Content = new TextBlock
            {
                Text = message,
                Margin = new Avalonia.Thickness(16),
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
            },
        };
        await dialog.ShowDialog(this);
    }
}
