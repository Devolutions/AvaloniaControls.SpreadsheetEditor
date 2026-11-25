namespace SpreadsheetEditor.Tests;

using System.Linq;
using System.Threading.Tasks;

using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

public sealed class MoreColorsTests
{
    [AvaloniaTheory]
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public async Task MoreColorsStagesCustomTextAndFillColorsUntilConfirmed(bool fill, bool accept, bool dark)
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[] { new[] { CellData.Plain("Formatted") with
            {
                Foreground = Colors.Red, Background = Colors.Yellow, Bold = true,
            } } },
        });
        Window window = new()
        {
            Content = editor, Width = 1100, Height = 600,
            RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light,
        };
        window.Show();
        try
        {
            await Frame(window);
            editor.FindControl<SpreadsheetGrid>("Grid")!.FocusForSelection();
            CellData original = editor.ViewModel.Rows[0].Cells[0].ToData();
            Button owner = editor.FindControl<Button>(fill ? "BackgroundButton" : "ForegroundButton")!;
            Flyout palette = Assert.IsType<Flyout>(owner.Flyout);
            palette.ShowAt(owner);
            await Frame(window);
            Button more = editor.FindControl<Button>(fill ? "MoreFillColorsButton" : "MoreTextColorsButton")!;
            Assert.Equal(EditorText.Get(EditorString.MoreColors), more.Content);
            more.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Frame(window);
            Flyout custom = Assert.IsType<Flyout>(owner.Flyout);
            Assert.NotSame(palette, custom);
            StackPanel content = Assert.IsType<StackPanel>(custom.Content);
            ColorView picker = Assert.Single(content.Children.OfType<ColorView>());
            Assert.Equal(fill ? Colors.Yellow : Colors.Red, picker.Color);
            Assert.False(picker.IsAlphaEnabled);
            Assert.NotEmpty(picker.GetVisualDescendants().OfType<ColorSpectrum>());
            Color chosen = Color.FromRgb(0x23, 0x67, 0x9A);
            picker.Color = chosen;
            await Frame(window);
            Assert.Equal(original, editor.ViewModel.Rows[0].Cells[0].ToData());
            Button action = content.GetVisualDescendants().OfType<Button>().Single(button =>
                Equals(button.Content, EditorText.Get(accept ? EditorString.OK : EditorString.Cancel)));
            action.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Frame(window);
            Assert.Same(palette, owner.Flyout);
            CellData expected = accept
                ? fill ? original with { Background = chosen } : original with { Foreground = chosen }
                : original;
            Assert.Equal(expected, editor.ViewModel.Rows[0].Cells[0].ToData());
            Assert.Equal(accept, editor.CanUndo);
            if (accept)
            {
                editor.Undo();
                Assert.Equal(original, editor.ViewModel.Rows[0].Cells[0].ToData());
            }
        }
        finally
        {
            editor.FindControl<Button>("ForegroundButton")!.Flyout?.Hide();
            editor.FindControl<Button>("BackgroundButton")!.Flyout?.Hide();
            editor.NewSheet(1, 1);
            window.Close();
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        }
    }

    [AvaloniaFact]
    public void ReadOnlyDisablesBothMoreColorsCommands()
    {
        EditorControl editor = new() { IsReadOnly = true };
        foreach (string name in new[] { "MoreTextColorsButton", "MoreFillColorsButton" })
        {
            Button button = editor.FindControl<Button>(name)!;
            Assert.False(button.IsEnabled);
            button.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        }
        editor.NewSheet(1, 1);
    }

    private static async Task Frame(Window window)
    {
        window.UpdateLayout();
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Render);
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
    }
}
