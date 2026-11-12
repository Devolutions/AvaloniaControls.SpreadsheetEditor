using System.Linq;
using System.Threading.Tasks;

using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;

using SpreadsheetEditor.Controls;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Models;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;

using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class HeaderMenuTests
{
    [AvaloniaTheory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public async Task RightClickOpensHeaderMenu(bool column, bool divider)
    {
        EditorControl editor = new();
        editor.LoadSheet(new SheetData
        {
            Cells = new[]
            {
                new[] { CellData.Plain("Name"), CellData.Plain("Value") },
                new[] { CellData.Plain("Zulu"), CellData.Plain("2") },
                new[] { CellData.Plain("Alpha"), CellData.Plain("1") },
            },
        });
        Window window = new() { Content = editor, Width = 800, Height = 500 };
        window.Show();
        try
        {
            await Frame(window);
            SpreadsheetGrid grid = editor.FindControl<SpreadsheetGrid>("Grid")!;
            Border target = grid.GetVisualDescendants().OfType<Border>().Single(border =>
                Equals(border.Tag, 1) && (divider
                    ? border.Name == (column ? "ColumnResizeHandle" : "RowResizeHandle")
                    : border.Child is TextBlock text && text.Text == (column ? "B" : "2")));
            MenuFlyout menu = Assert.IsType<MenuFlyout>(target.ContextFlyout);
            Point point = target.TranslatePoint(new Point(target.Bounds.Width / 2, target.Bounds.Height / 2), window)!.Value;
            window.MouseDown(point, MouseButton.Right);
            await Frame(window);
            window.MouseUp(point, MouseButton.Right);
            await Frame(window);
            Assert.True(menu.IsOpen);
            MenuFlyoutPresenter presenter = Assert.IsType<MenuFlyoutPresenter>(menu.Popup.Child);
            presenter.UpdateLayout();
            Assert.Equal(menu.Items.Count, presenter.ItemCount);
            Assert.Contains(presenter.GetVisualDescendants().OfType<MenuItem>(), item =>
                item.Bounds.Height > 0 && Equals(item.Header,
                    EditorText.Get(column ? EditorString.SortAscending : EditorString.InsertRowAbove)));
            menu.Hide();
        }
        finally
        {
            editor.NewSheet(1, 1);
            window.Close();
            await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
        }
    }

    private static async Task Frame(Window window)
    {
        window.UpdateLayout();
        window.CaptureRenderedFrame();
        await Dispatcher.UIThread.InvokeAsync(static () => { }, DispatcherPriority.SystemIdle);
        window.UpdateLayout();
        window.CaptureRenderedFrame();
    }
}
