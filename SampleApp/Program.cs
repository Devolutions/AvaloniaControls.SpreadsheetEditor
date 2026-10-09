namespace SampleApp;

using Avalonia;
using System;
using System.Globalization;
using System.Linq;

internal class Program
{
    // Initialization code. Don't use any Avalonia, third-party APIs or any
    // SynchronizationContext-reliant code before AppMain is called: things aren't initialized
    // yet and stuff might break.
    [STAThread]
    public static void Main(string[] args)
    {
        // Demo: --culture=fr shows the translated sample/editor labels before XAML is created.
        string? cultureOption = args.FirstOrDefault(arg => arg.StartsWith("--culture=", StringComparison.OrdinalIgnoreCase));
        if (cultureOption is not null)
        {
            CultureInfo culture = CultureInfo.GetCultureInfo(cultureOption["--culture=".Length..]);
            CultureInfo.DefaultThreadCurrentCulture = culture;
            CultureInfo.DefaultThreadCurrentUICulture = culture;
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;
            args = args.Where(arg => !string.Equals(arg, cultureOption, StringComparison.OrdinalIgnoreCase)).ToArray();
        }

        BuildAvaloniaApp().StartWithClassicDesktopLifetime(args);
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
#if DEBUG
            .WithDeveloperTools()
#endif
            .WithInterFont()
            .LogToTrace();
}
