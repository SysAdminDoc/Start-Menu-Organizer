# Research - Start Menu Organizer

## Executive Summary
Start Menu Organizer is a Windows 10/11 PowerShell 5.1 WPF utility for scanning Start Menu folders, detecting junk/broken/duplicate shortcuts, flattening folders, categorizing shortcuts, renaming entries, and backing up/restoring the menu. Its strongest shape is a focused local cleaner with a rich GUI and no runtime dependencies, but the highest-value direction is trust: convert every destructive operation into an auditable plan with persistent rollback, safe restore semantics, async workers, tests, and a packaged release. Top opportunities: fail-closed restore, persistent action journal, transaction preview/plan/export, async scan/action execution, real settings persistence, protected-folder/category rules, Pester/PSScriptAnalyzer coverage, installable artifact, and accessibility/i18n cleanup.

## Product Map
- Core workflows: scan user/system Start Menu folders; filter/search/sort entries; delete junk/broken/duplicate shortcuts; flatten or categorize folders; batch rename shortcuts; backup and restore Start Menu contents.
- User personas: Windows power users cleaning installer clutter; IT admins standardizing local machines; users recovering from stale uninstall/help/update shortcuts; maintainers shipping a small no-dependency tool.
- Platforms and distribution: Windows 10/11, Windows PowerShell 5.1+, WPF/XAML in `StartMenuOrganizerPro.ps1`; README currently ships source-script execution only, no signed package or release artifact.
- Key integrations and data flows: `%APPDATA%\Microsoft\Windows\Start Menu\Programs`, `%ProgramData%\Microsoft\Windows\Start Menu\Programs`, WScript.Shell COM shortcut parsing, `%LOCALAPPDATA%\StartMenuOrganizerPro\Backups`, JSON config export/import.

## Competitive Landscape
- Start Menu Helper: closest OSS peer with installer generation, startup/background cleaning, explicit backup warnings, folder/name/file-type rules, duplicate/broken-link deletion, and setup scripts. Learn packaged install and durable options; avoid unattended recurring cleanup until rollback is stronger.
- windows-shortcut-organizer: uses a three-phase scan/classify/organize pipeline with editable JSON artifacts, dry-run, idempotency, desktop plus Start Menu sources, and priority-ordered rules. Learn inspectable plan files and deterministic idempotent operations; avoid CLI-only UX as the primary experience.
- ClearWinStart: offers preview mode, dry-run, config wizard, file logging with rotation, preserved system folders, validation-only mode, tests, type hints, and a Python API. Learn protected folder defaults, logging, and testable core separation; avoid emoji/Unicode UI text for PowerShell 5.1 compatibility.
- Win11Debloat: demonstrates that PowerShell Windows-maintenance tools benefit from CLI parameters, applying settings to other users/default profiles, reversible changes, wiki docs, and no-install usage. Learn script parameterization and profile targeting; avoid broad OS tweaking that dilutes this repo's Start Menu focus.
- Open-Shell Menu: mature Start Menu replacement with releases, localization assets, ARM warnings, discussions, and large community signal. Learn release discipline, compatibility warnings, and localization pipeline; avoid shell replacement and explorer hooking.
- ExplorerPatcher: emphasizes installer/uninstaller/update flows, architecture-specific builds, elevated setup, and clear recovery/uninstall paths. Learn install/update/recovery clarity; avoid modifying Explorer behavior.
- Start11/StartAllBack/Start Menu X: commercial tools compete on polished layout control, grouping, search, backup/import/export, and Windows 11 restoration. Learn UX polish and backup/import value; avoid paid-shell-level customization creep.
- PowerToys Run: adjacent launcher with plugin architecture and searchable workflows. Learn extensible matching/search patterns; avoid turning this cleaner into a launcher.

## Security, Privacy, and Reliability
- `StartMenuOrganizerPro.ps1:2048` and `StartMenuOrganizerPro.ps1:2054` delete current Start Menu contents before copying backup files; a partial copy failure can leave the menu empty. Restore should stage, validate, and swap/copy with rollback.
- `StartMenuOrganizerPro.ps1:34` defines `UndoFile`, but undo is only in-memory (`StartMenuOrganizerPro.ps1:80`, `StartMenuOrganizerPro.ps1:1306`); deleted-item recovery disappears after process exit and temp cleanup.
- Bulk actions mutate files directly from UI event handlers (`StartMenuOrganizerPro.ps1:2254` through `StartMenuOrganizerPro.ps1:2265`) and scan synchronously (`StartMenuOrganizerPro.ps1:1106`), so large Start Menus can freeze the UI.
- System-folder handling is only admin-skip logic (`StartMenuOrganizerPro.ps1:1320`, `StartMenuOrganizerPro.ps1:1476`, `StartMenuOrganizerPro.ps1:1594`) and lacks a preserved-folder allowlist like Accessories, Administrative Tools, Startup, and Windows PowerShell.
- `Test-ShortcutBroken` only treats `.exe` targets as candidates (`StartMenuOrganizerPro.ps1:984`), so broken document/appref-ms/url/folder shortcuts are missed and some installer targets are under-classified.
- `Set-Content -Encoding UTF8` on config export (`StartMenuOrganizerPro.ps1:2080`) adds a BOM in Windows PowerShell 5.1; acceptable for JSON but inconsistent with project memory guidance for PS 5.1 file writes.
- PSScriptAnalyzer 1.25.0 reports state-changing functions without `ShouldProcess`, unused variables at `StartMenuOrganizerPro.ps1:2029` and `StartMenuOrganizerPro.ps1:2356`, and an automatic-variable assignment risk for `$sender` at `StartMenuOrganizerPro.ps1:2409`.

## Architecture Assessment
- Split the single 2,468-line script into script module/core functions plus thin WPF launcher so scan/classify/plan/execute can be tested without opening a window.
- Introduce a transaction model: scan returns inventory, classifier returns proposed operations, executor applies an operation list, journal persists before/after paths and backup copies.
- Replace repeated direct `Remove-Item`/`Move-Item`/`Rename-Item` calls with one guarded file-operation layer using `-LiteralPath`, same-root validation, collision policy, and structured result objects.
- Add async WPF workers using a background PowerShell/runspace plus dispatcher updates for scan and bulk actions; current progress bar does not prevent UI thread work.
- Persist settings automatically to `$Config.ConfigFile`; current import/export is manual and app startup does not load saved config.
- Add Pester fixtures that build temporary Start Menu trees and `.lnk` files, then exercise duplicate, broken, junk, flatten, rename, backup, restore, and config import/export behavior.
- Add README and release docs for artifact installation, rollback, admin behavior, and compatibility limits; no changelog, ignore file, working-notes file, or packaging files exist, but this research pass updates only `RESEARCH.md` and `ROADMAP.md`.

## Rejected Ideas
- Full Start Menu replacement: Open-Shell, ExplorerPatcher, Start11, and StartAllBack already own that category; replacing the shell contradicts this repo's cleanup-only purpose.
- Background auto-clean on startup before a durable journal: Start Menu Helper supports it, but unattended deletes are too risky while rollback is session-only.
- Package-manager integration as a first milestone: winget/chocolatey/scoop shortcut cleanup signals are real, but the project first needs safe transactions and tests.
- Launcher/search replacement: PowerToys Run covers extensible launch/search; Start Menu Organizer should remain a maintenance tool.
- Cloud sync or multi-user fleet service: Win11Debloat-style profile targeting is useful later, but cloud/fleet sync is unnecessary for a local Start Menu cleaner.

## Sources
Direct OSS peers:
- https://github.com/jarikmarwede/Start-Menu-Helper
- https://github.com/qwerd53/windows-shortcut-organizer
- https://github.com/RuanCH0924/ClearWinStart
- https://github.com/Raphire/Win11Debloat
- https://github.com/Open-Shell/Open-Shell-Menu
- https://github.com/valinet/ExplorerPatcher

Commercial and adjacent tools:
- https://www.stardock.com/products/start11/
- https://www.startallback.com/
- https://www.startmenux.com/
- https://learn.microsoft.com/en-us/windows/powertoys/run

Platform and standards:
- https://learn.microsoft.com/en-us/windows/configuration/start/layout
- https://learn.microsoft.com/en-us/windows/win32/shell/links
- https://learn.microsoft.com/en-us/windows/win32/api/shobjidl_core/nn-shobjidl_core-ishelllinkw
- https://learn.microsoft.com/en-us/windows/win32/shell/knownfolderid
- https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies

Testing, analysis, and packaging:
- https://github.com/PowerShell/PSScriptAnalyzer
- https://github.com/pester/Pester
- https://learn.microsoft.com/en-us/powershell/scripting/dev-cross-plat/create-standard-library-binary-module
- https://jrsoftware.org/isinfo.php
- https://pyinstaller.org/en/stable/

Community and package-manager signals:
- https://github.com/microsoft/winget-cli
- https://github.com/chocolatey/choco
- https://github.com/ScoopInstaller/Scoop
- https://github.com/ValveSoftware/steam-for-linux
- https://github.com/Heroic-Games-Launcher/HeroicGamesLauncher

## Open Questions
- Should the first packaged artifact be a signed `.ps1` plus launcher shortcut, a PS2EXE/PyInstaller-style executable, or an installer built with Inno Setup?
- Should system Start Menu changes require launching elevated up front, or should the app keep user-scope functionality available and gate only system-scope operations?
