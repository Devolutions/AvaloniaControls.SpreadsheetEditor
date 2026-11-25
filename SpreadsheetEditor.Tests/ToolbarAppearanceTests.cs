using System;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Headless.XUnit;
using Avalonia.Media;
using Avalonia.Styling;
using Avalonia.Threading;
using Avalonia.VisualTree;
using SpreadsheetEditor.Controls.Models;
using EditorControl = SpreadsheetEditor.Controls.SpreadsheetEditor;
using Xunit;

namespace SpreadsheetEditor.Tests;

public sealed class ToolbarAppearanceTests
{
    [AvaloniaTheory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task ReadOnlyToolbarDimsIconsTextAndSwatchesConsistently(bool dark, bool customButtonTheme)
    {
        Application app = Application.Current!;
        bool hadButtonTheme = app.Resources.ContainsKey(typeof(Button));
        object? previousButtonTheme = hadButtonTheme ? app.Resources[typeof(Button)] : null;
        if (customButtonTheme)
        {
            // Hosts can replace Button's theme while leaving ToggleButton on Fluent's fallback.
            // Explicitly colored icon content must still follow the editor's disabled appearance.
            ControlTheme hostTheme = new()
            {
                TargetType = typeof(Button),
                BasedOn = (ControlTheme)app.FindResource(typeof(Button))!,
            };
            hostTheme.Setters.Add(new Setter(Button.ForegroundProperty, Brushes.Magenta));
            hostTheme.Setters.Add(new Setter(Button.OpacityProperty, 0.8));
            app.Resources[typeof(Button)] = hostTheme;
        }
        EditorControl editor = new();
        editor.LoadSheet(new SheetData { Cells = new[] { new[] { CellData.Plain("Formatted") with
            { Bold = true, Foreground = Colors.Red, Background = Colors.Yellow } } } });
        Button hostButton = new() { Content = "Host button" };
        Grid content = new() { RowDefinitions = new RowDefinitions("*,Auto") };
        Grid.SetRow(hostButton, 1);
        content.Children.Add(editor); content.Children.Add(hostButton);
        Window window = new()
        {
            Content = content, Width = 1100, Height = 450,
            RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light,
        };
        window.Show();
        try
        {
            await Frame(window);
            if (customButtonTheme) Assert.Equal(0.8, hostButton.Opacity);
            Button italic = editor.FindControl<Button>("ItalicButton")!;
            PathIcon italicText = Assert.IsType<PathIcon>(italic.Content);
            Color enabledText = Ink(italicText);
            PathIcon fillIcon = Assert.Single(editor.FindControl<Button>("BackgroundButton")!.GetVisualDescendants().OfType<PathIcon>());
            Assert.Equal(enabledText, Ink(fillIcon));
            editor.IsReadOnly = true;
            await Frame(window);
            Button bold = editor.FindControl<Button>("BoldButton")!;
            TextBlock boldText = Assert.IsType<TextBlock>(bold.Content);
            Color disabledText = Ink(boldText);
            Assert.Equal(disabledText, Ink(fillIcon));
            Assert.NotEqual(enabledText, Ink(italicText));
            foreach (string name in new[]
            {
                "UndoButton", "RedoButton", "CutButton", "PasteButton", "ForegroundButton", "BackgroundButton",
                "BoldButton", "ItalicButton", "UnderlineButton", "StrikethroughButton",
                "AlignLeftButton", "AlignCenterButton", "AlignRightButton", "SymbolButton",
            })
            {
                Button button = editor.FindControl<Button>(name)!;
                Assert.False(button.IsEnabled, name);
                if (button.Content is PathIcon icon)
                    Assert.Equal(disabledText, Ink(icon));
                else if (button.Content is TextBlock caption)
                    Assert.Equal(disabledText, Ink(caption));
            }
            Assert.Equal(disabledText.A, Ink(editor.FindControl<Border>("ForegroundSwatch")!).A);
            Assert.Equal(disabledText.A, Ink(editor.FindControl<Border>("BackgroundSwatch")!).A);
            foreach (string name in new[] { "CopyButton", "PrintButton", "OverflowButton" })
                Assert.NotEqual(disabledText, Ink(Assert.IsType<PathIcon>(editor.FindControl<Button>(name)!.Content)));

            editor.IsReadOnly = false;
            await Frame(window);
            foreach (string name in new[] { "CutButton", "PasteButton", "ForegroundButton", "BackgroundButton", "BoldButton", "ItalicButton", "AlignLeftButton", "SymbolButton" })
            {
                Button button = editor.FindControl<Button>(name)!;
                Assert.True(button.IsEnabled, name);
            }
            Assert.Equal(enabledText, Ink(italicText));
            Assert.Equal(enabledText, Ink(fillIcon));
            Assert.Equal(enabledText, Ink(Assert.IsType<PathIcon>(editor.FindControl<Button>("PasteButton")!.Content)));
            Assert.True(editor.FindControl<Avalonia.Controls.Primitives.ToggleButton>("BoldButton")!.IsChecked);
            PathIcon alignment = Assert.IsType<PathIcon>(editor.FindControl<Button>("AlignLeftButton")!.Content);
            Assert.Equal(Ink(boldText), Ink(alignment));
        }
        finally
        {
            editor.NewSheet(1, 1); window.Close();
            if (customButtonTheme)
            {
                if (hadButtonTheme) app.Resources[typeof(Button)] = previousButtonTheme;
                else app.Resources.Remove(typeof(Button));
            }
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        }
    }

    private static Color Ink(Control glyph)
    {
        IBrush? brush = glyph switch
        {
            PathIcon icon => icon.Foreground, TextBlock text => text.Foreground, Border border => border.Background, _ => null,
        };
        ISolidColorBrush foreground = Assert.IsAssignableFrom<ISolidColorBrush>(brush);
        double opacity = foreground.Opacity;
        for (Visual? visual = glyph; visual is not null; visual = visual.GetVisualParent()) opacity *= visual.Opacity;
        Color color = foreground.Color;
        return Color.FromArgb((byte)Math.Round(color.A * opacity), color.R, color.G, color.B);
    }

    private static async Task Frame(Window window)
    {
        window.UpdateLayout(); window.CaptureRenderedFrame();
        await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.SystemIdle);
        window.UpdateLayout(); window.CaptureRenderedFrame();
    }
}
