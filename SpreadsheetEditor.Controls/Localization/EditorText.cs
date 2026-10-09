namespace SpreadsheetEditor.Controls.Localization;

using System;
using System.Globalization;
using System.Resources;
using Avalonia.Markup.Xaml;

// ResourceManager uses CurrentUICulture, just like RDM's UIResources/MsgResources.
// A missing translation falls back to the neutral English .resx.
public static class EditorText
{
    private static readonly ResourceManager Resources = new(
        "SpreadsheetEditor.Controls.Localization.SpreadsheetStrings", typeof(EditorText).Assembly);

    public static string Get(EditorString key) =>
        Resources.GetString(key.ToString(), ResourceCulture()) ?? key.ToString();

    public static string Format(EditorString key, params object?[] args) =>
        string.Format(CultureInfo.CurrentCulture, Get(key), args);

    // .NET's zh-TW parent is zh-Hant, not RDM's zh-Hant-TW resource name.
    private static CultureInfo ResourceCulture()
    {
        CultureInfo culture = CultureInfo.CurrentUICulture;
        string name = culture.Name;
        return name.StartsWith("zh-Hant", StringComparison.OrdinalIgnoreCase) ||
               name is "zh-TW" or "zh-HK" or "zh-MO"
            ? CultureInfo.GetCultureInfo("zh-Hant-TW")
            : culture;
    }
}

// Static XAML lookup matching RDM's UIRes usage. Recreate the control after changing UI culture.
public sealed class Loc(EditorString key) : MarkupExtension
{
    public override object ProvideValue(IServiceProvider serviceProvider) => EditorText.Get(key);

    public string ProvideTypedValue() => EditorText.Get(key);
}
