# Contributing

Thanks for your interest in contributing to **AvaloniaControls.SpreadsheetEditor**!

## Before you start

> [!IMPORTANT]
> This control depends on the commercial **Avalonia Accelerate** packages
> (`Avalonia.Controls.RichTextEditor`, `Avalonia.Controls.Charts`). You need your **own valid Avalonia
> UI license key** to build or run it — see <https://avaloniaui.net/>. The project itself is MIT-licensed,
> but it cannot be built without a license key.

## Prerequisites

- [.NET SDK 10.0](https://dotnet.microsoft.com/) or later.
- A valid Avalonia UI (Accelerate) license key.

## Building locally

Provide your license key via a **git-ignored `.env`** file at the repository root (never commit it):

```
AVALONIA_LICENSE_KEY=your-key-here
```

(Alternatively set an `AVALONIA_LICENSE_KEY` environment variable.) Then:

```bash
dotnet build SpreadSheetEditor.sln                   # build everything
dotnet run --project SampleApp                       # launch the sample app
dotnet pack SpreadsheetEditor.Controls -c Release    # build the NuGet package
```

## Project layout

| Path | Description |
|------|-------------|
| `SpreadsheetEditor.Controls/` | The reusable control library (the published NuGet package). |
| `SampleApp/` | A sample desktop host app used to exercise the control. |
| `Common.props`, `Directory.Build.props` | Shared build settings, dependency versions, license handling. |

## Coding standards

- Code style is enforced by the root [`.editorconfig`](.editorconfig) (Devolutions RDM C# conventions):
  explicit types (no `var`), `this.` qualification, Allman braces, `camelCase` fields, `AvaloniaList`
  over `ObservableCollection`, and usings inside the file-scoped namespace.
- Keep changes focused; match the style and comment density of the surrounding code.
- The library targets `net10.0` and must stay **cross-platform** (Windows/macOS/Linux) and pure-managed —
  avoid platform-specific dependencies (platform-specific code must be guarded and degrade gracefully).

## Verifying changes

There is no separate unit-test project; verify with a clean build and by exercising the change in the
sample app:

```bash
dotnet build SpreadSheetEditor.sln -c Release
dotnet run --project SampleApp
```

## Pull requests

1. Fork the repo and create a topic branch.
2. Make your change with a clear, focused commit history.
3. Ensure `dotnet build` is clean (no new warnings) and the sample app still runs.
4. **Never commit your license key** (`.env` is git-ignored — keep it that way).
5. Open a PR describing the change and how you tested it.

## Reporting bugs / requesting features

Use the [issue templates](.github/ISSUE_TEMPLATE/). For security issues, do **not** open a public issue —
see [SECURITY.md](SECURITY.md).
