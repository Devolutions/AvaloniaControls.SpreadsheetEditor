namespace SpreadsheetEditor.Tests;

using System.Linq;
using System.Threading.Tasks;

using Avalonia.Controls;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

public sealed class ReadOnlyTests
{
    [AvaloniaFact]
    public void ReadOnlyDisablesEditingControlsAndRestoresThemWhenCleared()
    {
        EditorControl editor = new();
        string[] editingControls =
        {
            "CutButton", "PasteButton", "FontFamilyBox", "FontSizeBox", "ForegroundButton",
            "BackgroundButton", "BoldButton", "ItalicButton", "UnderlineButton", "StrikethroughButton",
            "AlignLeftButton", "AlignCenterButton", "AlignRightButton", "SymbolButton", "AddSheetButton",
            "ReplaceButton", "ReplaceAllButton",
        };
        EditorString[] editingMenuItems =
        {
            EditorString.InsertChartMenu, EditorString.AutoSumMenu, EditorString.Functions,
            EditorString.AddRow, EditorString.AddColumn,
        };

        editor.IsReadOnly = true;

        foreach (string name in editingControls)
        {
            Assert.False(editor.FindControl<Control>(name)!.IsEnabled, name);
        }
        foreach (EditorString header in editingMenuItems)
        {
            Assert.False(OverflowItem(editor, header).IsEnabled, header.ToString());
        }
        foreach (string name in new[] { "CopyButton", "PrintButton", "OverflowButton", "ZoomInButton", "ZoomOutButton", "ZoomResetButton", "FindNextButton", "FindPreviousButton" })
        {
            Assert.True(editor.FindControl<Control>(name)!.IsEnabled, name);
        }
        Assert.True(OverflowItem(editor, EditorString.FindReplaceMenu).IsEnabled);
        Assert.True(OverflowItem(editor, EditorString.PrintMenu).IsEnabled);

        editor.IsReadOnly = false;

        foreach (string name in editingControls)
        {
            Assert.True(editor.FindControl<Control>(name)!.IsEnabled, name);
        }
        foreach (EditorString header in editingMenuItems)
        {
            Assert.True(OverflowItem(editor, header).IsEnabled, header.ToString());
        }
    }

    [AvaloniaFact]
    public void ReadOnlyBlocksAddingSheetsThroughTheApiAndButton()
    {
        EditorControl editor = new() { IsReadOnly = true };
        int changes = 0;
        editor.ContentChanged += (_, _) => changes++;

        editor.AddSheet();
        editor.FindControl<Button>("AddSheetButton")!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));

        Assert.Single(editor.Sheets);
        Assert.Equal(0, changes);

        editor.IsReadOnly = false;
        editor.AddSheet();
        Assert.Equal(2, editor.Sheets.Count);
        Assert.Equal(1, changes);
    }

    [AvaloniaTheory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReadOnlyBlocksHistoryWithoutDiscardingIt(bool redo)
    {
        EditorControl editor = new();
        CellAddress address = new(0, 0);
        editor.ViewModel.SetCell(address, "original");
        editor.ViewModel.Checkpoint();
        editor.ViewModel.SetCell(address, "edited");
        if (redo)
        {
            editor.Undo();
        }
        editor.IsReadOnly = true;
        int changes = 0;
        editor.ContentChanged += (_, _) => changes++;

        editor.Undo();
        editor.Redo();

        Assert.Equal(redo ? "original" : "edited", editor.ViewModel.GetCell(address)!.Text);
        Assert.Equal(0, changes);
        Assert.False(editor.FindControl<Button>("UndoButton")!.IsEnabled);
        Assert.False(editor.FindControl<Button>("RedoButton")!.IsEnabled);

        editor.IsReadOnly = false;
        Assert.Equal(!redo, editor.FindControl<Button>("UndoButton")!.IsEnabled);
        Assert.Equal(redo, editor.FindControl<Button>("RedoButton")!.IsEnabled);
        if (redo)
        {
            editor.Redo();
        }
        else
        {
            editor.Undo();
        }
        Assert.Equal(redo ? "edited" : "original", editor.ViewModel.GetCell(address)!.Text);
    }

    [AvaloniaFact]
    public void ReadOnlyBlocksSheetChangesAndStillAllowsSwitchingSheets()
    {
        EditorControl editor = new();
        editor.AddSheet();
        editor.IsReadOnly = true;
        StackPanel tabs = editor.FindControl<StackPanel>("SheetTabs")!;
        Button first = Assert.IsType<Button>(tabs.Children[0]);
        MenuFlyout menu = Assert.IsType<MenuFlyout>(first.ContextFlyout);
        MenuItem rename = Assert.IsType<MenuItem>(menu.Items[0]);
        MenuItem delete = Assert.IsType<MenuItem>(menu.Items[1]);
        Assert.False(rename.IsEnabled);
        Assert.False(delete.IsEnabled);

        rename.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
        delete.RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));

        Assert.IsType<TextBlock>(first.Content);
        Assert.Equal(2, editor.Sheets.Count);
        first.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        Assert.Same(editor.Sheets[0], editor.ViewModel);
        Assert.True(editor.FindControl<SpreadsheetGrid>("Grid")!.ReadOnly);

        editor.IsReadOnly = false;
        menu = Assert.IsType<MenuFlyout>(((Button)tabs.Children[0]).ContextFlyout);
        Assert.All(menu.Items.OfType<MenuItem>(), item => Assert.True(item.IsEnabled));
    }

    [AvaloniaFact]
    public void BecomingReadOnlyDuringRenameDoesNotCommitTheName()
    {
        EditorControl editor = new();
        Button tab = Assert.IsType<Button>(editor.FindControl<StackPanel>("SheetTabs")!.Children[0]);
        MenuFlyout menu = Assert.IsType<MenuFlyout>(tab.ContextFlyout);
        ((MenuItem)menu.Items[0]!).RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
        TextBox name = Assert.IsType<TextBox>(tab.Content);
        name.Text = "Changed name";
        string original = editor.Sheets[0].Name;

        editor.IsReadOnly = true;
        name.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });

        Assert.Equal(original, editor.Sheets[0].Name);
    }

    [AvaloniaTheory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public async Task ChartInsertionRespectsReadOnlyBeforeAndAfterTheDialog(bool initiallyReadOnly, bool becomeReadOnly)
    {
        EditorControl editor = new() { IsReadOnly = initiallyReadOnly };
        Window window = new() { Content = editor, Width = 900, Height = 650 };
        window.Show();
        try
        {
            OverflowItem(editor, EditorString.InsertChartMenu).RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
            if (initiallyReadOnly)
            {
                Assert.Empty(window.OwnedWindows);
            }
            else
            {
                Window dialog = Assert.Single(window.OwnedWindows);
                editor.IsReadOnly = becomeReadOnly;
                dialog.Close(new ChartBuilderResult(ChartKind.Line, "A1:B3", "Chart", false, true, true, false,
                    AxisNumberFormat.Number, null, null, null));
                await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
            }
            Assert.Equal(initiallyReadOnly || becomeReadOnly ? 0 : 1, editor.ViewModel.Charts.Count);
        }
        finally
        {
            foreach (Window owned in window.OwnedWindows.ToArray())
            {
                owned.Close();
            }
            editor.NewSheet(1, 1);
            window.Close();
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
        }
    }

    private static MenuItem OverflowItem(EditorControl editor, EditorString header)
    {
        MenuFlyout menu = Assert.IsType<MenuFlyout>(editor.FindControl<Button>("OverflowButton")!.Flyout);
        return menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, EditorText.Get(header)));
    }

    [AvaloniaFact]
    public async Task ReadOnlyDisablesContextMenuEditsWhileKeepingCopyAvailable()
    {
        EditorControl editor = new();
        Window window = new() { Content = editor, Width = 900, Height = 650 };
        window.Show();
        try
        {
            window.UpdateLayout();
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            foreach (bool readOnly in new[] { true, false })
            {
                editor.IsReadOnly = readOnly;
                Control[] targets = grid.GetVisualDescendants().OfType<Control>()
                    .Where(control => control.ContextFlyout is MenuFlyout menu &&
                        menu.Items.OfType<MenuItem>().Any(item =>
                            Equals(item.Header, EditorText.Get(EditorString.Cut)) ||
                            Equals(item.Header, EditorText.Get(EditorString.InsertRowAbove)) ||
                            Equals(item.Header, EditorText.Get(EditorString.SortAscending))))
                    .ToArray();
                Assert.NotEmpty(targets);
                foreach (Control target in targets)
                {
                    MenuFlyout menu = (MenuFlyout)target.ContextFlyout!;
                    menu.ShowAt(target);
                    await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
                    foreach (MenuItem item in menu.Items.OfType<MenuItem>())
                    {
                        bool expected = !readOnly || Equals(item.Header, EditorText.Get(EditorString.Copy));
                        Assert.True(expected == item.IsEnabled, $"{item.Header}: expected enabled={expected}, actual={item.IsEnabled}");
                    }
                    menu.Hide();
                }
            }
        }
        finally
        {
            window.Close();
        }
    }

    [AvaloniaTheory]
    [InlineData(EditorString.InsertComment)]
    [InlineData(EditorString.HyperlinkMenu)]
    [InlineData(EditorString.FormatCellsMenu)]
    public async Task BecomingReadOnlyWhileCellDialogIsOpenPreventsChanges(EditorString command)
    {
        EditorControl editor = new();
        editor.NewSheet(2, 2);
        Window window = new() { Content = editor, Width = 900, Height = 650 };
        window.Show();
        try
        {
            window.UpdateLayout();
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            MenuFlyout menu = grid.GetVisualDescendants().OfType<Control>()
                .Select(control => control.ContextFlyout).OfType<MenuFlyout>()
                .Single(flyout => flyout.Items.OfType<MenuItem>()
                    .Any(item => Equals(item.Header, EditorText.Get(command))));
            menu.Items.OfType<MenuItem>().Single(item => Equals(item.Header, EditorText.Get(command)))
                .RaiseEvent(new RoutedEventArgs(MenuItem.ClickEvent));
            Window dialog = Assert.Single(window.OwnedWindows);
            dialog.UpdateLayout();
            TextBox input = dialog.GetVisualDescendants().OfType<TextBox>().First();
            input.Text = command == EditorString.HyperlinkMenu ? "https://example.com" : "Changed";
            editor.IsReadOnly = true;
            int changes = 0;
            editor.ContentChanged += (_, _) => changes++;

            dialog.GetVisualDescendants().OfType<Button>().Single(button => button.IsDefault)
                .RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);

            CellData cell = editor.GetCellData()[0][0];
            Assert.Null(cell.Comment);
            Assert.Null(cell.Hyperlink);
            Assert.Null(cell.NumberFormat);
            Assert.Equal(0, changes);
        }
        finally
        {
            foreach (Window owned in window.OwnedWindows.ToArray())
            {
                owned.Close();
            }
            editor.NewSheet(1, 1);
            window.Close();
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
        }
    }
}
