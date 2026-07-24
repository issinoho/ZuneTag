# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ZuneTag ("Zune Meta Tag Editor" by The Drunken Bakery) is a Windows Forms desktop app (.NET Framework 4.8) for editing WMV metadata tags so video files display correctly as Movies, TV Shows, Music Videos, or generic Videos on a Zune. It's a single-form (`Form1`) WinForms app; `ZuneTag.Tests` covers the handful of pure-logic classes.

## Build

This is a classic .NET Framework (non-SDK-style) solution — build on Windows with MSBuild or Visual Studio, not `dotnet build`.

```
msbuild ZuneTag.sln /t:Restore,Build /p:Configuration=Release /p:Platform=x86
```

- Solution: `ZuneTag.sln` — contains `ZuneTag.csproj` (the WinForms app), `WMFSDKWrapper\ManagedWMFSDKWrapper.csproj` (COM interop wrapper, must build first; `ZuneTag.csproj` has a project reference to it), and `ZuneTag.Tests\ZuneTag.Tests.csproj` (xUnit tests, project-references `ZuneTag.csproj`).
- Target framework: `net48`. Platforms: `AnyCPU`, `x64`, `x86` (the NSIS installer packages the `x86\Release` output).
- NuGet packages for the main app restore via `packages.config` (old-style, not `PackageReference`) into `packages\`; `ZuneTag.Tests` is SDK-style and uses `PackageReference`.
- Run tests: `dotnet test ZuneTag.Tests\ZuneTag.Tests.csproj` (or via Visual Studio Test Explorer). Requires Windows — net48 test execution needs the Windows desktop runtime.
- Linting: `StyleCop.Analyzers` runs on all three projects (`ZuneTag.csproj` and `WMFSDKWrapper\ManagedWMFSDKWrapper.csproj` via `packages.config` + `<Analyzer>` items, `ZuneTag.Tests` via `PackageReference`), configured by the shared `stylecop.json` at the repo root. Rules run at their default severity (warning) — they surface in the build output and in Visual Studio's Error List but do not fail the build.
- For anything not covered by `ZuneTag.Tests` (i.e. most of `Form1.cs`, which is tightly coupled to WinForms UI state and native COM interop), verify changes by building and running the app manually (Windows only; requires the Windows Media Format SDK's `WMVCore.dll` to be registered/present, since the app P/Invokes into it).
- Installer script: `Installer\ZuneTag.nsi` (NSIS), packages `bin\x86\Release\*`.

## Architecture

**`Form1.cs`** is the entire application — a large partial class (~1800 lines) that owns UI state, WMV tag I/O, and TMDB search, all in one file. There's no layering into services/controllers; event handlers call straight into helper methods on `Form1` itself.

Key flows:

1. **WMV attribute editing** goes through native Windows Media Format SDK COM interop, wrapped in the separate `WMFSDKWrapper` project (`WMFSDKWrapper\WMFSDKFunctions.cs`):
   - `WMFSDKFunctions.WMCreateEditor` creates an `IWMMetadataEditor`, opened against a file path.
   - That editor is cast to `IWMHeaderInfo3` to enumerate/get/set/add/delete individual attributes (`GetAttributeByIndexEx`, `SetAttribute`, `AddAttribute`, `SetPicAttribute` for cover art, etc.).
   - `Form1` wraps these calls in its own helpers (`ShowAttributes3`, `ModifyAttrib`, `AddAttrib`, `AttribExists`, `EditorOpenFile`) which open/close the COM editor per call rather than holding it open.
   - Attribute values round-trip as raw byte arrays; `HexEncoding.cs` and `PrintAttribute` handle hex/UTF-16 (with BOM detection) decoding for display.
   - `Attribute.cs` is the app's own model class wrapping one WMV attribute (index/name/value/`WMT_ATTR_DATATYPE`) — distinct from `System.Attribute`; note the `using WMFSDKWrapper;` + custom class both named `Attribute` in this codebase, which shadows the base class type.

2. **Media type classification**: video files are typed by two GUID-valued attributes, `WM/MediaClassPrimaryID` and `WM/MediaClassSecondaryID`. `Form1` hardcodes the GUID string constants for Video/Movie/Music/TV (`TypeVideo`, `TypeMovie`, `TypeMusic`, `TypeTv`) and switches its UI panels (`gbMovie`/`gbMusic`/`gbTV`/`gbVideo`) based on which pair is present. `InspectFile()` re-reads all attributes from disk and re-derives this classification and the type-specific text fields every time a file is loaded or an attribute is modified.

3. **Metadata lookup via TMDB**: `cmdSearchTmdb_Click` triggers `backgroundWorker1` to call `TMDbClient.SearchMovieAsync` off the UI thread. Results are wrapped in `TMDbSearchResult.cs` (holds a `TMDbLib` `SearchMovie` plus derived genre/director/URL). Selecting a result lazily fetches extended movie/credits data and cover art (`ShowCoverArt`, via `WebClient` against `Settings.Default.PosterBase`). `cmdCopyResult_Click` copies the selected search result's fields into the currently-active media-type panel's text boxes — it does not write to the file directly; the per-type Save buttons (`cmdMovieSave_Click`, `cmdTVSave_Click`, etc.) do that via `EditAttribute`.
   - The TMDB API key lives in `Properties/Settings.settings` (`Settings.Default.APIkey`) as an application-scoped setting.

4. **Thumbnail generation**: `RegisterNewMediaFileAsync` uses `Xabe.FFmpeg` to snapshot a still frame from the loaded WMV (`FFmpeg.Conversions.FromSnippet.Snapshot`) into a temp PNG for preview; `FFmpegDownloader.GetLatestVersion` is invoked at form startup to ensure an FFmpeg binary is available.

When adding new WMV tag fields, follow the existing per-media-type pattern: add the attribute name to `AddMissingAttributes()` (so it's created if absent), read it in the matching `Load*Attributes()` method, and write it in the matching `cmd*Save_Click` handler via `EditAttribute`/`ModifyAttrib`.

## Testing

`ZuneTag.Tests` covers `HexEncoding`, `Attribute`, and `TMDbSearchResult` — the only classes with no WinForms/COM dependency. `Attribute` and `TMDbSearchResult` are `internal`; visibility to the test assembly is granted via `[assembly: InternalsVisibleTo("ZuneTag.Tests")]` in `Properties/AssemblyInfo.cs`. `Form1` itself isn't covered: it's a single WinForms class whose constructor does COM/network/FFmpeg setup, so testing it would require extracting logic out of it first.
