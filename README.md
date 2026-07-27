<p align="center">
  <img src="docs/icon.png" alt="ZuneTag icon" width="96">
</p>

# ZuneTag

**ZuneTag** ("Zune Meta Tag Editor" by The Drunken Bakery) is a Windows desktop tool for editing the metadata tags on `.wmv` files so they show up correctly as **Movies**, **TV Shows**, **Music Videos**, or generic **Videos** when synced to a Zune.

The Zune classifies video content using a pair of GUID-valued attributes baked into the WMV file (`WM/MediaClassPrimaryID` / `WM/MediaClassSecondaryID`). Most tools that produce or re-encode WMV files don't set these correctly, so files land in the wrong category (or none at all) on the device. ZuneTag lets you inspect and fix a file's raw attributes directly, and fills in the surrounding metadata (title, description, genre, year, cast/director, cover art, etc.) to match.

## Features

- Reads and writes WMV metadata attributes directly via the Windows Media Format SDK, including the Zune-specific media-classification GUIDs.
- Dedicated panels for Movie, TV Show, Music Video, and generic Video metadata, matching what the Zune expects for each type.
- Built-in search against [The Movie Database (TMDB)](https://www.themoviedb.org/) to pull in title, genre, director, description, and cover art, then copy it straight into the tag fields.
- Generates a thumbnail preview from the video file itself so you can confirm you're tagging the right one.
- Logs diagnostic output to `%TEMP%\ZuneTag.log` for troubleshooting.

## Tech stack

- **.NET Framework 4.8**, Windows Forms (single-window desktop app).
- **Windows Media Format SDK** (`WMVCore.dll`) via a small COM-interop wrapper project, for all attribute reading/writing.
- **[TMDbLib](https://github.com/LordMike/TMDbLib)** for TMDB search and metadata.
- **[Xabe.FFmpeg](https://github.com/tomaszzmuda/Xabe.FFmpeg)** for thumbnail generation (FFmpeg itself is downloaded automatically on first run).
- **[Fody](https://github.com/Fody/Fody) / [Costura.Fody](https://github.com/Fody/Costura)** to bundle all managed dependencies into a single `ZuneTag.exe` for release builds.
- **xUnit** for the test suite covering the app's pure-logic classes.
- Built and published via **GitHub Actions** on `windows-latest`; packaged for install via an **NSIS** installer script.

## Downloading and running it

Grab the latest `ZuneTag.exe` from the [Releases page](https://github.com/issinoho/ZuneTag/releases) — it's a single portable executable with all managed dependencies bundled in, so there's nothing else to install alongside it.

**Prerequisites:**

- **Windows** (7 SP1 or later).
- **[.NET Framework 4.8 runtime](https://dotnet.microsoft.com/en-us/download/dotnet-framework/net48)** installed. Most current Windows installs already have this; if `ZuneTag.exe` won't launch, this is the first thing to check.
- **The Windows Media Format SDK's `WMVCore.dll`**, registered and present on the system. This ships with Windows Media Player and is present on most consumer Windows installs by default; it may be missing on server SKUs or "N"/"KN" editions of Windows that omit media features, in which case installing the relevant Windows Media Feature Pack resolves it.
- **An internet connection on first run** — the app downloads an FFmpeg binary automatically (for thumbnail generation) and needs connectivity for TMDB search. No API key setup is required; a TMDB key is already bundled in the app.

No installation is required for the standalone EXE — just run it. An NSIS-based installer (`Installer/ZuneTag.nsi`) is also available if you'd prefer a proper Start Menu install; see `CLAUDE.md` for how to build it.

## Building from source

See [`CLAUDE.md`](CLAUDE.md) for build instructions, project layout, and architecture notes.

## License

ZuneTag is freeware — see [`License.txt`](License.txt) for the full terms (personal and internal corporate use is fine; it can't be bundled with or sold as part of a commercial product).
