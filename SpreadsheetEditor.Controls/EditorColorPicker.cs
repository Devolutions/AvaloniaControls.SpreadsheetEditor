namespace SpreadsheetEditor.Controls;

using System;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Markup.Xaml.Styling;
using Avalonia.Media;

internal static class EditorColorPicker
{
    internal static ColorView CreateView(Color color) => Configure(new ColorView(), color);

    internal static ColorPicker CreateDropDown(Color color) => Configure(new ColorPicker(), color);

    private static T Configure<T>(T picker, Color color) where T : ColorView
    {
        // Match RDM's Markdown toolbar, including the neutral colours before the Fluent palette.
        picker.HsvColor = color == Colors.Black ? new HsvColor(1, 0, 1, 0) : color.ToHsv();
        picker.ColorSpectrumComponents = ColorSpectrumComponents.HueValue;
        picker.IsAlphaEnabled = false;
        picker.IsAlphaVisible = false;
        picker.Palette = new FluentColorPalette();
        picker.SelectedIndex = (int)ColorViewTab.Spectrum;
        picker.PaletteColors = new[] { Colors.Black, Colors.DimGray, Colors.Gray, Colors.Silver, Colors.WhiteSmoke, Colors.White }
            .Concat(picker.PaletteColors ?? Array.Empty<Color>()).ToArray();
        picker.Styles.Add(new StyleInclude(new Uri("avares://SpreadsheetEditor.Controls/"))
        {
            Source = new Uri("avares://Avalonia.Controls.ColorPicker/Themes/Fluent/Fluent.xaml"),
        });
        return picker;
    }
}
