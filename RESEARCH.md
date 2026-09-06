# Research: Start Menu Organizer

## Executive Summary
Start Menu Organizer is a Windows PowerShell 5.1/WPF utility for auditing, reviewing, backing up, and safely reorganizing user, system, and profile Start Menu shortcut folders. Verified: the current app has the right shape for a local trust-first cleaner with background scans, editable operation plans, rollback journals, protected system folders, structured JSONL logs, accessibility and localization metadata, profile targeting, and a zip release package. The highest-value direction is to harden the filesystem and shortcut-inspection trust boundary before adding broader automation. Top opportunities include reparse-point protection, full `.lnk` and `.url` metadata risk classification, Known Folder API path resolution, atomic state writes, release signing, human-readable audit exports, package-manager shortcut resurrection handling, configurable rule presets, runtime accessibility checks, and administrator relaunch.

## Product Map
- Core workflows: scan Start Menu scopes, classify shortcuts/folders, preview/edit an operation plan, execute guarded moves/deletes/restores, inspect logs/history.
- User personas: Windows power user cleaning installer clutter, technician standardizing a machine, admin preparing default/profile menus, accessibility-conscious user needing predictable UI labels.
- Platforms and distribution: Windows 10/11 desktop; Windows PowerShell 5.1; WPF; local zip package from `tools/Build-Release.ps1`; no runtime package dependency beyond built-in Windows assemblies and `WScript.Shell`.
- Key integrations and data flows: Start Menu filesystem roots in `%APPDATA%`, `%PROGRAMDATA%`, profile/default-user folders; `.lnk`, `.url`, `.appref-ms` shortcuts; `%LOCALAPPDATA%\StartMenuOrganizerPro` config, backup, undo journal, localization override, and JSONL logs.

## Competitive Landscape
- Start Menu Helper: does continuous cleaning, backup prompts, setup install, and folder flattening well. Learn from its install/setup and configurable cleanup modes; avoid unattended mutation until Start Menu Organizer has stronger provenance, reparse, and shortcut-risk safeguards.
- Start Menu Cleaner: small focused cleaner with executable release, CLI options, logging, and shortcut edit/create support. Learn from the CLI/reporting path; avoid packaging patterns that increase AV friction.
- Start-Menu-Manager: uses WPF, MSI packaging, JSON shortcut definitions, console builder, custom icons, and uninstall cleanup. Learn from portable shortcut schemas and MSI-style install polish. Avoid turning this cleaner into a shortcut-authoring suite before metadata preservation is reliable.
- Open-Shell, ExplorerPatcher, StartAllBack, and Windhawk: prove demand for deep Start/taskbar customization and compatibility tracking after Windows updates. Learn from explicit compatibility/release notes and modular customization; intentionally avoid shell replacement, Explorer injection, and system hooking.
- Start11: makes backup/restore, role/kiosk layouts, tabs/groups, and deployment scripting paid enterprise features. Learn from exportable configurations and fleet handoff; avoid paywall-style feature sprawl unrelated to local cleanup.
- Start Menu X: virtual groups avoid real-folder drift after installs/upgrades/uninstalls. Learn from non-mutating previews; avoid replacing filesystem cleanup with a parallel launcher database.
- PowerToys and Win11Debloat: show Windows utilities win trust with privacy posture, multiple install channels, explicit presets, and local-first execution. Learn from install/distribution clarity and preset discipline; avoid broad OS-debloat creep.
- TileIconifier: demonstrates shortcut argument/icon metadata bugs in real Start workflows. Learn from preserving target arguments, icons, and metadata; avoid visual tile theming as a core roadmap item.

## Security, Privacy, and Reliability
- Verified risk: `StartMenuOrganizerPro.ps1` validates string paths with `Test-PathWithinRoot`, but recursive scan/copy/delete/restore paths do not explicitly block or neutralize reparse points before `Get-ChildItem -Recurse`, `Copy-Item -Recurse`, or `Remove-Item -Recurse` (`StartMenuOrganizerPro.ps1:2105`, `StartMenuOrganizerPro.ps1:2225`, `StartMenuOrganizerPro.ps1:2883`, `StartMenuOrganizerPro.ps1:4060`, `StartMenuOrganizerPro.ps1:4107`, `StartMenuOrganizerPro.ps1:4113`).
- Verified risk: shortcut inspection reads only the target path for `.lnk` files and only the `URL=` line for `.url` files (`StartMenuOrganizerPro.ps1:1956`, `StartMenuOrganizerPro.ps1:1984`, `StartMenuOrganizerPro.ps1:2903`). Shell Link and AppUserModelID metadata include arguments, working directory, icon, description, property store, and AppUserModelID; recent LNK/Internet Shortcut CVEs and malware research make hidden arguments and crafted shortcuts a trust boundary.
- Verified gap: release packaging uses `-ExecutionPolicy Bypass` in the generated Start Menu shortcut and emits no hash manifest or signing status (`tools/Build-Release.ps1:67`). PowerShell signing/execution-policy docs support adding Authenticode signing when a certificate is available and always publishing SHA256 verification data.
- Verified gap: config, journal, and log writes use direct whole-file writes; invalid JSON handling exists, but writes are not atomic with a last-known-good backup (`StartMenuOrganizerPro.ps1:1273`, `StartMenuOrganizerPro.ps1:2276`, `StartMenuOrganizerPro.ps1:2562`, `StartMenuOrganizerPro.ps1:4379`).
- Missing guardrails: suspicious shortcut classification, reparse-point reporting, release hash verification, provenance for shortcuts recreated by package managers, and recovery tests that simulate partial state-file writes.
- Recovery and rollback needs: keep current backup/journal model, but add atomic writes, reportable skipped-risk items, and explicit rollback evidence for every guarded operation.

## Architecture Assessment
- Boundary improvement: split shortcut metadata extraction/risk classification out of UI scan flow into testable functions near `Get-ShortcutTarget` and `Test-ShortcutBroken`; expose structured fields to the DataGrid, JSONL log, plan JSON, and report export.
- Boundary improvement: centralize filesystem traversal and mutation safety in one reparse-aware helper used by scan, backup, delete, move, rename, and restore paths; current guards are strong for string roots but not enough for link traversal.
- Refactor candidate: replace hard-coded Start Menu/profile suffix construction in config and `Get-ProfileProgramsPath` with Known Folder API resolution plus safe fallbacks (`StartMenuOrganizerPro.ps1:32`, `StartMenuOrganizerPro.ps1:34`, `StartMenuOrganizerPro.ps1:1810`).
- Refactor candidate: move rule/category/junk/protected-folder matching into an ordered rule engine with conflict diagnostics; current data tables are useful but not externally inspectable or preset-friendly (`StartMenuOrganizerPro.ps1:58`, `StartMenuOrganizerPro.ps1:72`, `StartMenuOrganizerPro.ps1:84`).
- Test gap: `tests\Run-Tests.ps1` runs parser, Pester, and ScriptAnalyzer checks but has no coverage threshold, runtime WPF smoke, screenshot/visual-tree validation, or synthetic `.lnk` metadata fixtures.
- Documentation gap: README covers setup and workflows, but not release verification, shortcut risk semantics, rule precedence, localization template validation, or administrator access behavior.

## Rejected Ideas
- Full shell replacement or Explorer/taskbar hooking: Open-Shell, ExplorerPatcher, StartAllBack, and Windhawk cover this space; it contradicts the local cleaner model and raises stability/security risk.
- Always-on auto-cleaning as the next step: Start Menu Helper shows demand, but package updates recreate shortcuts and can conflict with user intent; provenance and safer rollback should land first.
- Broad Windows debloat/privacy/AI toggles: Win11Debloat and similar utilities already own this; it would dilute Start Menu-specific cleanup and testing.
- Tile/icon theming and custom launcher skinning: TileIconifier and Start11 already cover visual Start customization; this project should preserve metadata before creating new visual schemes.
- Cloud sync or hosted multi-user service: commercial Start tools offer fleet controls, but this repository is a local PowerShell/WPF app; exportable bundles and policy handoff fit better.
- Mobile support: no credible use case for a Windows Start Menu filesystem cleaner on mobile platforms.
- Plugin marketplace: Windhawk and PowerToys show modular ecosystems, but Start Menu Organizer needs rule presets and import/export before accepting third-party extension code.

## Sources
Direct OSS and adjacent projects:
- https://github.com/jarikmarwede/Start-Menu-Helper
- https://github.com/qwerty-w/start-menu-cleaner
- https://github.com/James231/Start-Menu-Manager
- https://github.com/Open-Shell/Open-Shell-Menu
- https://github.com/valinet/ExplorerPatcher
- https://github.com/Raphire/Win11Debloat
- https://github.com/microsoft/PowerToys
- https://github.com/Jonno12345/TileIconifier
- https://github.com/ramensoftware/windhawk

Commercial and closed-source:
- https://www.stardock.com/blog/523496/introducing-start11-v2
- https://www.startallback.com/
- https://www.startmenux.com/
- https://www.startmenux.com/help/Virtual%20Groups.html

Platform, standards, and security:
- https://learn.microsoft.com/en-us/windows/configuration/start/layout
- https://learn.microsoft.com/en-us/windows/win32/shell/knownfolderid
- https://learn.microsoft.com/en-us/windows/win32/api/shlobj_core/nf-shlobj_core-shgetknownfolderpath
- https://learn.microsoft.com/en-us/windows/win32/shell/links
- https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-shllink/a6c2f32d-2297-4727-bcd3-5d3669573bcb
- https://learn.microsoft.com/en-us/windows/win32/shell/appids
- https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_signing
- https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_execution_policies
- https://nvd.nist.gov/vuln/detail/CVE-2025-9491
- https://nvd.nist.gov/vuln/detail/CVE-2025-33053
- https://www.trendmicro.com/en_us/research/25/c/windows-shortcut-zero-day-exploit.html

Community, package managers, and testing:
- https://github.com/chocolatey/choco/issues/2016
- https://github.com/ScoopInstaller/Scoop/issues/4562
- https://www.reddit.com/r/Windows10/comments/zgtpkg/anyone_know_how_to_export_or_backup_a_windows/
- https://www.reddit.com/r/Windows11/comments/1rcrt5z/start_menu_remove_the_all_section_with_categories/
- https://github.com/PowerShell/PowerShell/issues/621
- https://pester.dev/docs/usage/code-coverage

## Open Questions
- Needs live validation: which Windows 10/11 builds in the supported range expose redirected or localized Start Menu Known Folder paths differently from the current hard-coded path fallbacks?
