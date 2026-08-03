# Changelog

## v0.14.0 - 2026-07-01

### Features
- Full shortcut metadata extraction (target, arguments, working dir, description, hotkey, icon) for `.lnk` and `.url` files.
- Risk classification flags: ScriptHost, NetworkTarget, HiddenExecution, LongArguments, WebTarget, ScriptTarget.
- Package-manager provenance detection (winget, Chocolatey, Scoop, MSIX, Installer).
- Reparse-point (junction/symlink) traversal blocking in scans and mutations.
- Atomic file writes with `.bak` fallback recovery for config and undo journal.
- Known Folder API resolution via SHGetKnownFolderPath with safe fallbacks.
- SHA256 manifest and optional Authenticode signing in release builds.
- Scan report export as CSV or JSON.
- Enterprise handoff export with deployment guidance.
- Virtual grouping preview (non-mutating category organization preview).
- Localization template export and completeness validation.
- Unified rule evaluation engine with preset import/export.
- Elevated relaunch button with capability status display.
- Pester code coverage reporting (JaCoCo format).
- winget, Chocolatey, and Scoop package metadata generation in builds.

### Fixes (audit pass)
- Fixed all action functions using `-Path` instead of `-LiteralPath` (wildcard characters in folder names could cause silent failures or wrong-path matches).
- Fixed `Apply-Filters` null-safety for items with null TargetPath/DisplayName/RelativePath.
- Fixed `Reset-Configuration` to actually reset categories (previously only reset junk patterns despite the dialog saying both).
- Fixed `Remove-AllJunk` and `Remove-BrokenShortcuts` silently skipping items filtered out of the DataGrid view.
- Fixed `Move-ToCategory` accepting null category name without validation.
- Fixed `NetworkTarget` risk flag false positive on any path containing double backslashes (now only flags UNC paths).
- Fixed `Flatten-SingleItemFolders` TOCTOU: now re-checks folder is empty after move before deleting.
- Fixed `Write-AtomicFile` leaving orphan `.tmp` files on failure.
- Fixed `Find-Replace-Names` using `-like` with user input containing wildcard characters.
- Fixed COM shortcut object leaks in `Get-ShortcutTarget`, `Get-ShortcutMetadata`, and all worker equivalents.
- Fixed elevation button using wrong script path when invoked from non-file scriptblocks.
- Fixed imported operation plans not validating paths against approved mutation roots at import time.
- Fixed missing `colRisk.Header` localization key.
- Fixed `btnElevate` missing from tab order.
- Fixed COM shortcut object leak in installer script.
- Added atomic-write backup-fallback recovery regression test.

## v0.13.0 - 2026-06-30

- Added Selected Profile and Default User scope options with explicit profile-root validation.
- Added profile/default-user scope metadata for scanning, guarded mutations, separate backup folders, and restore planning.
- Added admin gating for system, selected-profile, and default-user mutations through the guarded file-operation helper.
- Added profile/default-user regression coverage.

## v0.12.0 - 2026-06-28

- Added centralized UI string metadata with optional JSON localization overrides under `%LOCALAPPDATA%\StartMenuOrganizerPro\Localization`.
- Added automation names, help text, and tooltip backfill for core interactive controls.
- Added an explicit tab-order list for primary scan, action, plan, settings, and log controls.
- Raised muted text contrast and added contrast/localization regression coverage.

## v0.11.0 - 2026-06-28

- Added dated JSONL file logging under `%LOCALAPPDATA%\StartMenuOrganizerPro\Logs`.
- Added crash-log files for unhandled WPF dispatcher, AppDomain, and top-level runtime exceptions before the user-facing error message is shown.
- Added operation IDs to guarded operation results, journal items, and journal-related log entries.
- Added log rotation and structured logging regression coverage.

## v0.10.0 - 2026-06-28

- Added a local release-package builder that creates `StartMenuOrganizer-v0.10.0.zip`.
- Added current-user install and uninstall scripts with a Start Menu shortcut and execution-policy launch guidance.
- Added package regression coverage to verify the artifact name, included files, installer shortcut wiring, and uninstall path.
- Documented installable package usage, uninstall command, and admin behavior for system Start Menu changes.

## v0.9.0 - 2026-06-28

- Added a unified local test gate that runs parser checks, disposable-fixture regression scripts, Pester, and PSScriptAnalyzer.
- Added PSScriptAnalyzer settings with documented suppressions for WPF-local function naming and ShouldProcess noise.
- Added Pester smoke coverage for parser and version badge alignment.
- Fixed concrete analyzer findings for unused context-menu variables and automatic-variable event parameter naming.

## v0.8.0 - 2026-06-28

- Added versioned settings persistence for scope, junk patterns, protected folders, and category patterns.
- Loaded saved settings on startup and auto-saved scope and settings edits.
- Updated config import/export to use the same schema and recover from invalid JSON by moving bad files aside.
- Added settings persistence regression coverage.

## v0.7.0 - 2026-06-28

- Added built-in protected Start Menu folder detection for folders such as Startup, Windows Tools, Windows PowerShell, and Administrative Tools.
- Skipped protected items in delete, flatten, empty-folder cleanup, root move, category move, auto-organize, and batch rename paths.
- Expanded shortcut validation for `.url`, `.appref-ms`, file URLs, documents, folders, environment-variable paths, shell links, and app protocol targets.
- Added shortcut validation and protected-folder regression coverage.

## v0.6.0 - 2026-06-28

- Moved reviewed operation-plan execution into a cancelable dispatcher worker with progress updates.
- Added partial-journal finalization when a running plan is canceled.
- Kept synchronous plan execution available for the local regression harness.

## v0.5.0 - 2026-06-28

- Moved Start Menu scanning into a cancelable PowerShell runspace worker.
- Added dispatcher-marshaled scan completion so item binding, filtering, stats, status, and logs update on the UI thread.
- Added a header Cancel action for active background work.
- Added async scan wiring regression coverage.

## v0.4.0 - 2026-06-28

- Added editable transaction plans generated from Preview Mode for bulk delete, cleanup, move, organize, flatten, and rename actions.
- Added plan summaries with operation counts and before/after paths in the activity log.
- Added JSON plan export/import plus Execute/Clear controls that run the exact loaded plan through the guarded journaled operation path.
- Added operation-plan regression coverage for JSON round-trip and mixed delete/move/rename execution.

## v0.3.0 - 2026-06-28

- Centralized destructive delete, move, and rename operations behind one guarded file-operation helper.
- Added literal-path mutation, approved-root validation, same-root move enforcement, collision failure policy, preview results, structured operation results, and rollback backup registration.
- Routed journaled actions, undo moves/renames, and restore directory clearing through the guarded helper.
- Added guarded file-operation regression coverage for preview, cross-root rejection, collision handling, and rollback backups.

## v0.2.0 - 2026-06-28

- Added a persistent JSON undo journal with durable undo backups under `%LOCALAPPDATA%\StartMenuOrganizerPro`.
- Journaled successful and skipped/failed destructive delete, move, rename, flatten, restore, and cleanup operations with original path, new path, backup path, action type, timestamp, and result fields.
- Added startup journal loading so the last reversible action can be undone after restarting the app.
- Added a persistent undo regression test that deletes a temp shortcut, reloads the journal, and restores it through undo.

## v0.1.0 - 2026-06-28

- Added fail-closed restore planning that stages backup contents, validates restore sources, and preserves pre-restore rollback snapshots before live Start Menu entries are replaced.
- Added restore rollback-on-failure handling and warnings for skipped or partially failed restore scopes.
- Added a local restore safety regression test and README test instructions.
