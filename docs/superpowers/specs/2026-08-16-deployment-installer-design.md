# Deployment: Inno Setup Installer

## Context

eneBridge 2.0 has no packaging/distribution story yet — it's only ever been run via `dotnet run`
or a raw `dotnet build` output folder. It needs to reach a non-technical end user (an accounting
clerk) on a small, known number of machines, so a guided one-click setup matters more than a
minimal/portable footprint.

The app depends on two external prerequisites that must be installed as real Windows
components regardless of how the app itself is packaged:

- **Microsoft Access Database Engine** (`prerequisites/accessdatabaseengine_X64.exe`) — the
  OleDb/ACE dBASE IV driver `DbfExportService` writes through. This is a COM/ODBC driver; it
  cannot be made "portable," it must be installed.
- **.NET 8 Desktop Runtime** (`prerequisites/dotnet-runtime-8.0.23-win-x64.exe`) — required
  because the app is published framework-dependent (see below).

The original lost-source app (`reference/eneBridge`) was distributed as a flat,
framework-dependent folder drop with no installer — these two prerequisite installers already
existing in `prerequisites/` are carried over from that setup. `appsettings.json`'s shipped
defaults (`C:\Christine\source\data.xlsx`, `C:\emas\ch2025\acc\data\`) are leftover
dev/test values from that era and don't reflect any real packaging requirement — the GUI is fully
browse-driven, and `SettingsService` persists whatever the user picks to
`%AppData%\eneBridge\settings.json`, overriding the shipped default on every subsequent run.

## Goal

Produce a single `setup.exe` a non-technical user can double-click to get a fully working
install — app + both prerequisites — with no separate manual steps, and support painless
in-place upgrades when a new build needs to go out.

## Design

### Overall approach

Build the installer with **Inno Setup** (free, script-based, the standard tool for exactly this
pattern — bundle an app plus a couple of prerequisite EXEs, chain-install them conditionally, add
Start Menu/uninstall entries). The app is published **framework-dependent**
(`dotnet publish -c Release`, no `--self-contained`) rather than self-contained, since the
installer is already responsible for getting the .NET 8 Desktop Runtime onto the machine.

The Inno Setup script uses a **fixed `AppId` GUID**, generated once and baked into the script
permanently. This is what lets a later, newer `setup.exe` be recognized as an upgrade of the same
product rather than a separate parallel install (see Updates/patching below).

On launch, the installer checks the registry for each prerequisite's installed state and only
runs the corresponding bundled EXE (silently, via Inno Setup's `[Run]` section with a `Check:`
condition) if it isn't already present — so re-running the installer on a machine that already has
both prerequisites doesn't touch them again. The installer requires the standard Windows UAC
elevation prompt (needed to write to Program Files and to run the prerequisite installers) — this
is normal, expected behavior for this class of installer and requires no special handling.

### Install layout & shortcuts

- Installs to `%ProgramFiles%\eneBridge` (Inno Setup's standard 64-bit default).
- Start Menu shortcut created unconditionally; a Desktop shortcut is offered as a checkbox on the
  install wizard's "select additional tasks" page (Inno Setup's standard built-in page — no custom
  UI needed).
- A normal "Apps & Features" / "Programs and Features" uninstall entry is registered. Uninstalling
  removes the app's files and shortcuts but **leaves the Access Database Engine and .NET Desktop
  Runtime installed** — standard practice for shared system components, since other software on
  the machine may depend on them and safely detecting "is anything else using this" isn't
  practical for an uninstaller to do.
- User data — `SettingsService`'s `%AppData%\eneBridge\settings.json`, `RunHistoryService`'s
  `%AppData%\eneBridge\runHistory.json`, and `FileLogger`'s `%AppData%\eneBridge\logs\` — lives
  entirely outside the Program Files install folder and is untouched by install, upgrade, or
  uninstall.

### Updates / patching

No in-app update-check logic (explicitly out of scope — see below). Patching is: bump the version
number (in the `.iss` script, and ideally the WPF project's assembly version too, so both agree),
rebuild `setup.exe`, and hand the new installer to whoever runs it on the target machine(s).

Because the `AppId` is fixed, Inno Setup recognizes the new `setup.exe` as an upgrade of the
existing install: it detects the prior version, closes the running app if it's open, replaces the
application files in place, and leaves the same Start Menu shortcut and Add/Remove Programs entry
(now showing the new version). No manual uninstall step is required, and
`%AppData%\eneBridge\` (settings, run history, logs) survives automatically since it's outside the
install folder entirely.

### Pre-packaging cleanup

Before the first installer build, blank the placeholder values in
`src/eneBridge.Wpf/appsettings.json`:

```json
{
  "FilePaths": {
    "ExcelFilePath": "",
    "DbfFilePath": ""
  }
}
```

so a fresh install shows empty path fields rather than another user's old folder layout. (This is
safe precisely because `SettingsService` already overrides these with the user's last-used paths
from `%AppData%` once they've run the app once — see CLAUDE.md's Persistence section.)

### Build pipeline

A new `installer/` folder holds the Inno Setup script, `eneBridge.iss`. It references:

- The `dotnet publish` output folder (the compiled app + dependencies).
- The two existing files in `prerequisites/` (`accessdatabaseengine_X64.exe`,
  `dotnet-runtime-8.0.23-win-x64.exe`), bundled into the installer via Inno Setup's `[Files]`
  section and invoked conditionally from `[Run]`.

Build steps:

```
dotnet publish src/eneBridge.Wpf/eneBridge.Wpf.csproj -c Release -o installer/publish
iscc installer/eneBridge.iss
```

producing `installer/Output/setup.exe` (Inno Setup's default output location). This should be
documented as a new command in CLAUDE.md's Commands section once the script exists.

### Explicitly out of scope (YAGNI at this scale)

- **No auto-update mechanism.** Manual upgrade-in-place (above) is sufficient for a small, known
  set of machines with an identifiable person distributing new installers when needed. An
  in-app update-check would require hosting new installers somewhere reachable (a shared folder or
  URL) and update-check/download logic in the app — real added scope with no current need.
- **No code-signing certificate.** Windows SmartScreen may show an "Unknown Publisher" warning the
  first time `setup.exe` runs on a machine. Accepted for now, since this is internal distribution
  where the user already expects and trusts the installer they were handed; revisit if this
  becomes real friction.
- **No MSIX/Store packaging.** Unnecessary for direct internal distribution outside the Microsoft
  Store.

## Verification

- Run the built `setup.exe` on a clean VM/machine without the prerequisites installed; confirm
  both silently install, the app installs and launches, and Start Menu/Desktop shortcuts work.
- Run it again on that same machine (prerequisites now present); confirm it does *not* re-run
  either prerequisite installer, and still completes the app install/upgrade correctly.
- Bump the version and rebuild `setup.exe`; run it on a machine with the previous version
  installed and already-open. Confirm the running app is closed, files are upgraded in place, and
  `%AppData%\eneBridge\settings.json`/`runHistory.json` survive unchanged.
- Uninstall via "Apps & Features"; confirm the app and shortcuts are removed, while the Access
  Database Engine and .NET Desktop Runtime remain installed.
