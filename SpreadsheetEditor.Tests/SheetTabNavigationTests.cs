using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Interactivity;
using Avalonia.Threading;
using SpreadsheetEditor.Controls.Models;
using SpreadsheetEditor.Controls.Localization;
using SpreadsheetEditor.Controls.Xlsx;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;
using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class SheetTabNavigationTests
{
    [AvaloniaFact]
    public async Task HiddenFirstSheetOpensTheFirstVisibleSheetAndFinishesLoading()
    {
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Lookup", new SheetData { Cells = new[] { new[] { CellData.Plain("hidden") } }, Visibility = "hidden" }),
            ("Visible", new SheetData { Cells = new[] { new[] { CellData.Plain("visible data") } } }),
        });
        using MemoryStream saved = new(); editor.SaveXlsx(saved); saved.Position = 0;
        using (ClosedXML.Excel.XLWorkbook workbook = new(saved))
            Assert.True(workbook.Worksheet("Visible").TabActive);
        saved.Position = 0;
        try
        {
            await editor.OpenXlsxAsync(saved);
            Assert.Equal("Visible", editor.ViewModel.Name);
            Assert.Equal("visible data", editor.ViewModel.Rows[0].Cells[0].Text);
            Stopwatch timeout = Stopwatch.StartNew();
            while (editor.FindControl<SpreadsheetEditor.Controls.SpreadsheetGrid>("Grid")!.ReadOnly && timeout.Elapsed < TimeSpan.FromSeconds(10))
                await Task.Delay(10);
            Assert.False(editor.FindControl<SpreadsheetEditor.Controls.SpreadsheetGrid>("Grid")!.ReadOnly);
            editor.ViewModel.SetCell(new CellAddress(0, 0), "edited");
            Assert.Equal("edited", editor.ViewModel.Rows[0].Cells[0].Text);
            Assert.Single(editor.FindControl<StackPanel>("SheetTabs")!.Children);
            using MemoryStream resaved = new();
            await editor.SaveXlsxAsync(resaved);
            resaved.Position = 0;
            var reopened = XlsxFile.ReadWorkbook(resaved);
            Assert.Equal("hidden", reopened[0].Data.Visibility);
            Assert.Equal("hidden", reopened[0].Data.Cells[0][0].Text);
            Assert.Equal("edited", reopened[1].Data.Cells[0][0].Text);
        }
        finally { editor.NewSheet(1, 1); await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle); }
    }

    [AvaloniaFact]
    public void DeleteSheetMenuKeepsTheLastVisibleSheet()
    {
        EditorControl editor = new();
        editor.LoadWorkbook(new[]
        {
            ("Visible", new SheetData { Cells = new[] { new[] { CellData.Plain("data") } } }),
            ("Lookup", new SheetData { Cells = new[] { new[] { CellData.Plain("lookup") } }, Visibility = "hidden" }),
        });
        MenuItem DeleteItem(Button tab) => Assert.IsType<MenuFlyout>(tab.ContextFlyout).Items.OfType<MenuItem>()
            .Single(item => Equals(item.Header, EditorText.Get(EditorString.Delete)));
        StackPanel tabs = editor.FindControl<StackPanel>("SheetTabs")!;
        Assert.False(DeleteItem(Assert.IsType<Button>(Assert.Single(tabs.Children))).IsEnabled);
        editor.AddSheet();
        Assert.All(tabs.Children.OfType<Button>(), tab => Assert.True(DeleteItem(tab).IsEnabled));
        editor.IsReadOnly = true;
        Assert.All(tabs.Children.OfType<Button>(), tab => Assert.False(DeleteItem(tab).IsEnabled));
    }

    [AvaloniaFact]
    public async Task ArrowsScrollOverflowWithoutChangingTheActiveSheet()
    {
        EditorControl editor = CreateWorkbook(25);
        Window window = new() { Content = editor, Width = 750, Height = 450 }; window.Show();
        try
        {
            await Frame(window);
            ScrollViewer scroll = editor.FindControl<ScrollViewer>("SheetTabScroller")!;
            RepeatButton left = editor.FindControl<RepeatButton>("ScrollSheetTabsLeftButton")!;
            RepeatButton right = editor.FindControl<RepeatButton>("ScrollSheetTabsRightButton")!;
            Assert.True(scroll.Extent.Width > scroll.Viewport.Width);
            Assert.False(left.IsEnabled);
            Assert.True(right.IsEnabled);
            Assert.True(scroll.TranslatePoint(new Point(scroll.Bounds.Width, 0), window)!.Value.X < window.Width);
            var selected = editor.ViewModel;
            right.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Frame(window);
            Assert.True(scroll.Offset.X > 0);
            Assert.True(left.IsEnabled);
            Assert.Same(selected, editor.ViewModel);
            double offset = scroll.Offset.X;
            editor.IsReadOnly = true; // Rebuilding tab colours/menus must preserve a manual scroll.
            await Frame(window);
            Assert.Equal(offset, scroll.Offset.X, precision: 1);
            for (int i = 0; i < 30; i++) right.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Frame(window);
            Assert.False(right.IsEnabled);
            Assert.Equal(scroll.Extent.Width - scroll.Viewport.Width, scroll.Offset.X, precision: 1);
            for (int i = 0; i < 30; i++) left.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            await Frame(window);
            Assert.False(left.IsEnabled);
            Assert.True(right.IsEnabled);
            Assert.Equal(0, scroll.Offset.X);
            Assert.Same(selected, editor.ViewModel);
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task AddingASheetRevealsItAndReopeningResetsTheScroll()
    {
        EditorControl editor = CreateWorkbook(20);
        Window window = new() { Content = editor, Width = 750, Height = 450 }; window.Show();
        try
        {
            await Frame(window);
            editor.AddSheet();
            await Frame(window);
            ScrollViewer scroll = editor.FindControl<ScrollViewer>("SheetTabScroller")!;
            StackPanel tabs = editor.FindControl<StackPanel>("SheetTabs")!;
            var active = tabs.Children.OfType<Button>().Single(tab => Equals(tab.Tag, editor.ViewModel.Name));
            Assert.True(scroll.Offset.X > 0);
            Assert.True(active.Bounds.Left >= scroll.Offset.X - 1);
            Assert.True(active.Bounds.Right <= scroll.Offset.X + scroll.Viewport.Width + 1);
            editor.LoadWorkbook(Sheets(20));
            await Frame(window);
            Assert.Equal(0, scroll.Offset.X);
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task ArrowsUpdateWhenResizingAndHiddenSheetsStayOutOfTheStrip()
    {
        EditorControl editor = new();
        editor.LoadWorkbook(Sheets(4).Append(("Hidden", new SheetData { Cells = new[] { new[] { CellData.Plain("lookup") } }, Visibility = "hidden" })).ToArray());
        Window window = new() { Content = editor, Width = 1400, Height = 450 }; window.Show();
        try
        {
            await Frame(window);
            RepeatButton left = editor.FindControl<RepeatButton>("ScrollSheetTabsLeftButton")!;
            RepeatButton right = editor.FindControl<RepeatButton>("ScrollSheetTabsRightButton")!;
            Assert.Equal(4, editor.FindControl<StackPanel>("SheetTabs")!.Children.Count);
            Assert.False(left.IsEnabled);
            Assert.False(right.IsEnabled);
            window.Width = 460;
            await Frame(window);
            Assert.True(right.IsEnabled);
            right.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
            window.Width = 1400;
            await Frame(window);
            Assert.False(left.IsEnabled);
            Assert.False(right.IsEnabled);
        }
        finally { await Close(editor, window); }
    }

    [AvaloniaFact]
    public async Task InitialLayoutKeepsTheSelectedTabVisibleAfterRepeatedRebuilds()
    {
        EditorControl editor = CreateWorkbook(20);
        editor.AddSheet(); // Select the last sheet before the control has a viewport.
        Window window = new() { Content = editor, Width = 750, Height = 450 }; window.Show();
        try
        {
            await Frame(window);
            ScrollViewer scroll = editor.FindControl<ScrollViewer>("SheetTabScroller")!;
            Button active = editor.FindControl<StackPanel>("SheetTabs")!.Children.OfType<Button>().Single(tab => Equals(tab.Tag, editor.ViewModel.Name));
            Assert.True(scroll.Offset.X > 0);
            Assert.True(active.Bounds.Left >= scroll.Offset.X - 1);
            Assert.True(active.Bounds.Right <= scroll.Offset.X + scroll.Viewport.Width + 1);
        }
        finally { await Close(editor, window); }
    }

    private static (string Name, SheetData Data)[] Sheets(int count) => Enumerable.Range(1, count)
        .Select(i => ($"Worksheet {i:00}", new SheetData { Cells = new[] { new[] { CellData.Plain(i.ToString()) } } })).ToArray();

    private static EditorControl CreateWorkbook(int count)
    {
        EditorControl editor = new(); editor.LoadWorkbook(Sheets(count)); return editor;
    }

    private static async Task Frame(Window window)
    {
        window.UpdateLayout(); window.CaptureRenderedFrame();
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        window.UpdateLayout(); window.CaptureRenderedFrame();
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
    }

    private static async Task Close(EditorControl editor, Window window)
    {
        editor.NewSheet(1, 1); window.Close();
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
    }
}
