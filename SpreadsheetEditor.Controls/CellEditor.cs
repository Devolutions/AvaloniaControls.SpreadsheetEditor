namespace SpreadsheetEditor.Controls;

using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Controls.Documents.TextModel;
using Avalonia.Controls.Primitives;
using Avalonia.Layout;
using Avalonia.Media;
using global::SpreadsheetEditor.Controls.Models;
using global::SpreadsheetEditor.Controls.ViewModels;

// The in-cell editor: a single instance the grid moves over the active cell. Prefers the Pro
// RichTextEditor (per-character bold/italic/underline/color), and falls back to a plain TextBox
// (whole-cell formatting) when the Pro control isn't available — e.g. no license, or
// Avalonia.Headless where it doesn't initialize. The toolbar calls the Apply*/Toggle* methods,
// which route to the rich selection or the whole-cell flags accordingly.
public class CellEditor : Decorator
{
    private readonly RichTextEditor? rich;
    private readonly TextBox? plain;
    private CellViewModel? cell;
    private bool loading;      // suppress write-back while we (re)load the document/text
    private bool committing;   // re-entrancy guard while extracting runs back to the cell

    public CellEditor()
    {
        try
        {
            // Strip the editor's default chrome (border/corner/padding/background/min-size) so it
            // fills the cell flat instead of rendering as a small rounded floating box. Vertical text
            // centering is handled via the document's PagePadding in BuildDocument.
            this.rich = new RichTextEditor
            {
                AcceptsReturn = false,
                ShowToolbar = false,
                ShowPageBounds = false,
                ShowBlockAdorners = false,
                // Suppress the editor's floating selection flyout (a mini formatting popup with a
                // dropdown). All its formatting lives in the grid's top toolbar instead.
                ShowSelectionFlyout = false,
                HorizontalAlignment = HorizontalAlignment.Stretch,
                VerticalAlignment = VerticalAlignment.Stretch,
                BorderThickness = new(0),
                Padding = new(0),
                CornerRadius = new(0),
                Background = Brushes.Transparent,
                MinHeight = 0,
                MinWidth = 0,
                // Stripping the chrome also drops the default selection/caret brushes — set them so
                // click-drag text selection shows a blue highlight and the caret is visible.
                SelectionBrush = new SolidColorBrush(Avalonia.Media.Color.FromArgb(0x66, 0x1A, 0x73, 0xE8)),
                CaretBrush = Brushes.Black,
            };

            // A cell is a single line, but the document's vertical-centering padding can make the
            // content a hair taller than the cell, which makes the editor's inner ScrollViewer show a
            // vertical scrollbar — its up-arrow ("line up") button appears as a chevron at the cell's
            // right edge and obstructs the content. Disable both scrollbars so no such chrome renders.
            ScrollViewer.SetVerticalScrollBarVisibility(this.rich, ScrollBarVisibility.Disabled);
            ScrollViewer.SetHorizontalScrollBarVisibility(this.rich, ScrollBarVisibility.Disabled);

            this.rich.ContentChanged += (_, _) => this.CommitFromRich();
            this.rich.SelectionChanged += (_, _) => this.SelectionFormatChanged?.Invoke(this, EventArgs.Empty);
            this.Child = this.rich;
        }
        catch
        {
            // Pro editor unavailable — degrade to a plain editable TextBox (whole-cell formatting).
            this.rich = null;
            this.plain = new TextBox
            {
                BorderThickness = new(0),
                Background = Brushes.Transparent,
                Padding = new(4, 2),
                AcceptsReturn = false,
                VerticalContentAlignment = VerticalAlignment.Center,
            };
            this.plain.TextChanged += (_, _) =>
            {
                if (!this.loading && this.cell is not null)
                {
                    this.cell.Text = this.plain.Text ?? string.Empty;
                }
            };
            this.Child = this.plain;
        }
    }

    // True when the rich editor is active (per-character formatting available).
    public bool IsRich => this.rich is not null;

    // Raised when the rich editor's caret/selection changes (so the toolbar can re-sync).
    public event EventHandler? SelectionFormatChanged;

    // Points the editor at a cell and loads its content.
    public void Load(CellViewModel cell)
    {
        this.cell = cell;
        this.loading = true;
        try
        {
            if (this.rich is not null)
            {
                this.rich.Document = BuildDocument(cell);
            }
            else if (this.plain is not null)
            {
                this.plain.Text = cell.Text;
                this.plain.FontWeight = cell.FontWeight;
                this.plain.FontStyle = cell.FontStyle;
                this.plain.FontSize = cell.FontSize;
                this.plain.TextAlignment = cell.TextAlignment;
                this.plain.Foreground = cell.Foreground ?? Brushes.Black;
                if (!string.IsNullOrWhiteSpace(cell.FontFamily))
                {
                    this.plain.FontFamily = new FontFamily(cell.FontFamily);
                }
            }
        }
        finally
        {
            this.loading = false;
        }
    }

    public void FocusEditor(bool selectAll = false)
    {
        (this.rich as Control ?? this.plain)?.Focus();
        if (selectAll)
        {
            try { this.rich?.SelectAll(); } catch { }
            this.plain?.SelectAll();
        }
        else if (this.plain is not null)
        {
            this.plain.CaretIndex = this.plain.Text?.Length ?? 0;
        }
    }

    // ----- Formatting the in-cell selection (per-character; springloads for a caret) -----
    // Each method reads the current state and applies the explicit opposite, so results are
    // deterministic (no blind toggle that can desync from what's shown).

    public void ToggleBold()
    {
        if (this.rich is null)
        {
            if (this.cell is not null)
            {
                this.cell.Bold = !this.cell.Bold;
            }

            this.SyncPlain();
            return;
        }

        this.ApplyToSelection(RichTextElement.FontWeightProperty, this.CurrentBold() ? FontWeight.Normal : FontWeight.Bold);
    }

    public void ToggleItalic()
    {
        if (this.rich is null)
        {
            if (this.cell is not null)
            {
                this.cell.Italic = !this.cell.Italic;
            }

            this.SyncPlain();
            return;
        }

        this.ApplyToSelection(RichTextElement.FontStyleProperty, this.CurrentItalic() ? FontStyle.Normal : FontStyle.Italic);
    }

    public void ToggleUnderline()
    {
        if (this.rich is null)
        {
            if (this.cell is not null)
            {
                this.cell.Underline = !this.cell.Underline;
            }

            this.SyncPlain();
            return;
        }

        (bool underline, bool strike) = this.CurrentDecorationFlags();
        this.ApplyToSelection(RichInline.TextDecorationsProperty, BuildDecorations(!underline, strike));
    }

    public void ToggleStrikethrough()
    {
        if (this.rich is null)
        {
            if (this.cell is not null)
            {
                this.cell.Strikethrough = !this.cell.Strikethrough;
            }

            this.SyncPlain();
            return;
        }

        (bool underline, bool strike) = this.CurrentDecorationFlags();
        this.ApplyToSelection(RichInline.TextDecorationsProperty, BuildDecorations(underline, !strike));
    }

    public void ApplyForeground(Color color)
    {
        if (this.rich is null)
        {
            if (this.cell is not null)
            {
                this.cell.ForegroundColor = color;
            }

            this.SyncPlain();
            return;
        }

        this.ApplyToSelection(RichTextElement.ForegroundProperty, (IBrush)new SolidColorBrush(color));
    }

    // The current B/I/U/strikethrough state of the selection (for keeping the toolbar in sync).
    public (bool Bold, bool Italic, bool Underline, bool Strikethrough) GetSelectionFlags()
    {
        if (this.rich is null)
        {
            return (this.cell?.Bold ?? false, this.cell?.Italic ?? false, this.cell?.Underline ?? false, this.cell?.Strikethrough ?? false);
        }

        (bool underline, bool strike) = this.CurrentDecorationFlags();
        return (this.CurrentBold(), this.CurrentItalic(), underline, strike);
    }

    private bool CurrentBold() => (int)this.GetSelected(RichTextElement.FontWeightProperty, FontWeight.Normal) >= (int)FontWeight.SemiBold;

    private bool CurrentItalic() => this.GetSelected(RichTextElement.FontStyleProperty, FontStyle.Normal) == FontStyle.Italic;

    // Underline and strikethrough share the rich editor's TextDecorations property; read both at once.
    private (bool Underline, bool Strikethrough) CurrentDecorationFlags() =>
        DecorationFlags(this.GetSelected(RichInline.TextDecorationsProperty, (TextDecorationCollection?)null));

    // Builds a decorations collection (null when neither is set) for underline/strikethrough.
    private static TextDecorationCollection? BuildDecorations(bool underline, bool strikethrough)
    {
        if (!underline && !strikethrough)
        {
            return null;
        }

        TextDecorationCollection decorations = new();
        if (underline)
        {
            decorations.Add(new TextDecoration { Location = TextDecorationLocation.Underline });
        }

        if (strikethrough)
        {
            decorations.Add(new TextDecoration { Location = TextDecorationLocation.Strikethrough });
        }

        return decorations;
    }

    private static (bool Underline, bool Strikethrough) DecorationFlags(TextDecorationCollection? decorations)
    {
        if (decorations is null)
        {
            return (false, false);
        }

        bool underline = decorations.Any(d => d.Location == TextDecorationLocation.Underline);
        bool strikethrough = decorations.Any(d => d.Location == TextDecorationLocation.Strikethrough);
        return (underline, strikethrough);
    }

    private void SyncPlain()
    {
        if (this.plain is null || this.cell is null)
        {
            return;
        }

        this.plain.FontWeight = this.cell.FontWeight;
        this.plain.FontStyle = this.cell.FontStyle;
        this.plain.Foreground = this.cell.Foreground ?? Brushes.Black;
    }

    // ----- Rich document <-> CellRun conversion -----

    private static FlowDocument BuildDocument(CellViewModel cell)
    {
        Paragraph paragraph = new() { TextAlignment = cell.TextAlignment };
        foreach (CellRun r in cell.EditRuns)
        {
            if (string.IsNullOrEmpty(r.Text))
            {
                continue;
            }

            RichRun run = new(r.Text);
            if (r.Bold)
            {
                run.SetValue(RichTextElement.FontWeightProperty, FontWeight.Bold);
            }

            if (r.Italic)
            {
                run.SetValue(RichTextElement.FontStyleProperty, FontStyle.Italic);
            }

            if (r.Underline || r.Strikethrough)
            {
                run.TextDecorations = BuildDecorations(r.Underline, r.Strikethrough);
            }

            if (r.Foreground is { } fg)
            {
                run.SetValue(RichTextElement.ForegroundProperty, new SolidColorBrush(fg));
            }

            paragraph.Inlines.Add(run);
        }

        if (paragraph.Inlines.Count == 0)
        {
            paragraph.Inlines.Add(new RichRun(string.Empty));
        }

        // The editor host is inset 2px from the cell; pad another 2px so text lines up with the
        // display TextBlock's 4px left margin. Top padding vertically centers the single line.
        double hostHeight = Math.Max(22.0, cell.FontSize + 8) - 4; // cell row height minus the 2px insets
        double topPad = Math.Max(0, (hostHeight - cell.FontSize * 1.3) / 2);
        FlowDocument document = new() { FontSize = cell.FontSize, PagePadding = new(2, topPad, 0, 0) };
        if (!string.IsNullOrWhiteSpace(cell.FontFamily))
        {
            document.FontFamily = new FontFamily(cell.FontFamily);
        }

        document.Blocks.Add(paragraph);
        return document;
    }

    // Reads the rich document back into the cell as mixed runs (per-character formatting).
    private void CommitFromRich()
    {
        if (this.loading || this.committing || this.rich?.Document is not { } doc || this.cell is null)
        {
            return;
        }

        try
        {
            this.committing = true;
            string text = doc.ContentRange?.GetText() ?? string.Empty;
            List<CellRunDraft> runs = new();
            TextPointer? contentStart = doc.ContentStart;
            if (contentStart is null)
            {
                return;
            }

            for (int i = 0; i < text.Length; i++)
            {
                TextPointer? start = contentStart.GetPositionAtOffset(i, LogicalDirection.Forward);
                TextPointer? end = contentStart.GetPositionAtOffset(i + 1, LogicalDirection.Forward);
                if (start is null || end is null)
                {
                    break;
                }

                TextRange range = new(start, end);
                bool bold = false, italic = false, underline = false, strikethrough = false;
                Color? foreground = null;
                try { bold = (int)range.GetPropertyValue(RichTextElement.FontWeightProperty) >= (int)FontWeight.SemiBold; } catch { }
                try { italic = range.GetPropertyValue(RichTextElement.FontStyleProperty) == FontStyle.Italic; } catch { }
                try { (underline, strikethrough) = DecorationFlags(range.GetPropertyValue(RichInline.TextDecorationsProperty)); } catch { }
                try { foreground = ColorOf(range.GetPropertyValue(RichTextElement.ForegroundProperty)); } catch { }
                runs.Add(new CellRunDraft(text[i], bold, italic, underline, foreground, strikethrough));
            }

            // Coalesce merges adjacent same-format characters, so a uniformly-formatted cell yields a
            // single run. In that case store it as a plain cell (whole-cell flags, no explicit runs) so
            // the number format still applies on display; keep explicit runs only for mixed formatting.
            List<CellRun> coalesced = Coalesce(runs);
            if (coalesced.Count <= 1)
            {
                CellRun u = coalesced.Count == 1 ? coalesced[0] : new CellRun(string.Empty, false, false, false, null, null);
                this.cell.SetUniformContent(text, u.Bold, u.Italic, u.Underline, u.Strikethrough, u.Foreground);
            }
            else
            {
                this.cell.SetRichRuns(coalesced);
            }
        }
        catch
        {
            // Defensive: never let a parsing hiccup in the Pro editor crash typing.
        }
        finally
        {
            this.committing = false;
        }
    }

    private static Color? ColorOf(IBrush? brush)
    {
        if (brush is not SolidColorBrush s)
        {
            return null;
        }

        // Treat default black as "unset" so plain text stays unstyled.
        return s.Color == Colors.Black ? null : s.Color;
    }

    // Merges adjacent single-character drafts with identical formatting into runs.
    private static List<CellRun> Coalesce(List<CellRunDraft> drafts)
    {
        List<CellRun> runs = new();
        foreach (CellRunDraft d in drafts)
        {
            if (runs.Count > 0
                && runs[^1].Bold == d.Bold && runs[^1].Italic == d.Italic
                && runs[^1].Underline == d.Underline && runs[^1].Strikethrough == d.Strikethrough && runs[^1].Foreground == d.Foreground)
            {
                runs[^1] = runs[^1] with { Text = runs[^1].Text + d.Char };
            }
            else
            {
                runs.Add(new CellRun(d.Char.ToString(), d.Bold, d.Italic, d.Underline, d.Foreground, null, d.Strikethrough));
            }
        }

        return runs;
    }

    private readonly record struct CellRunDraft(char Char, bool Bold, bool Italic, bool Underline, Color? Foreground, bool Strikethrough = false);

    // ----- Rich selection helpers -----

#pragma warning disable CS8600, CS8601, CS8603, CS8604 // unconstrained T over nullable rich properties
    // Applies a value to the selection (or springloads it onto an empty caret).
    private void ApplyToSelection<T>(Avalonia.StyledProperty<T> property, T value)
    {
        TextSelection? sel = this.rich?.Selection;
        if (sel is null)
        {
            return;
        }

        try
        {
            if (sel.IsEmpty)
            {
                sel.SetSpringloadProperty(property, value);
            }
            else
            {
                sel.ApplyPropertyValue(property, value);
            }
        }
        catch { }
    }

    // Reads the selection's value for a property (springload value for an empty caret),
    // falling back when the selection is mixed/unset.
    private T GetSelected<T>(Avalonia.StyledProperty<T> property, T fallback)
    {
        TextSelection? sel = this.rich?.Selection;
        if (sel is null)
        {
            return fallback;
        }

        try
        {
            if (sel.IsEmpty)
            {
                return sel.TryGetSpringloadProperty(property, out T sv) ? sv : fallback;
            }

            return sel.TryGetPropertyValue(property, out T v) ? v : fallback;
        }
        catch { return fallback; }
    }
#pragma warning restore CS8600, CS8601, CS8603, CS8604
}
