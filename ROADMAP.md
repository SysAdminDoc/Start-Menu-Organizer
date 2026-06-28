# Roadmap

## Research-Driven Additions

- [ ] P0 - Add persistent undo and operation journal
  Why: `$Config.UndoFile` exists but undo is only in memory and deletion recovery uses temp files that may disappear after exit.
  Evidence: `StartMenuOrganizerPro.ps1:34`, `StartMenuOrganizerPro.ps1:80`, `StartMenuOrganizerPro.ps1:1306`, ClearWinStart logging model.
  Touches: `StartMenuOrganizerPro.ps1` undo, delete, move, rename, backup storage.
  Acceptance: Every destructive operation writes a JSON journal with original path, new path, backup path, action type, timestamp, and result; undo after restart restores the last reversible action.
  Complexity: M

- [ ] P0 - Centralize guarded file operations
  Why: Destructive calls are scattered across actions with mixed error handling and wildcard paths.
  Evidence: `StartMenuOrganizerPro.ps1:1331`, `StartMenuOrganizerPro.ps1:1523`, `StartMenuOrganizerPro.ps1:1651`, `StartMenuOrganizerPro.ps1:1792`, PSScriptAnalyzer `PSUseShouldProcessForStateChangingFunctions`.
  Touches: `StartMenuOrganizerPro.ps1` delete/move/rename/restore helpers.
  Acceptance: All filesystem mutations go through one helper using `-LiteralPath`, same-root checks, collision policy, structured results, preview support, and rollback registration.
  Complexity: L

- [ ] P1 - Replace direct actions with editable transaction plans
  Why: Direct peers make risky shortcut changes safer with dry-run/editable plan artifacts before execution.
  Evidence: windows-shortcut-organizer scan/classify/organize pipeline, ClearWinStart preview/dry-run, current `chkPreviewMode`.
  Touches: `Refresh-Items`, action functions, preview log, optional plan JSON export/import.
  Acceptance: Bulk actions can generate an operation plan, show counts and before/after paths, export/import JSON, and execute the exact reviewed plan.
  Complexity: L

- [ ] P1 - Move scan and bulk actions off the UI thread
  Why: Recursive scans and bulk moves/deletes run from WPF event handlers and can freeze the window on large menus.
  Evidence: `StartMenuOrganizerPro.ps1:1106`, `StartMenuOrganizerPro.ps1:2254` through `StartMenuOrganizerPro.ps1:2265`, PowerShell stack async GUI convention.
  Touches: scan/action dispatch, progress/status updates, cancellation handling.
  Acceptance: Scan and bulk actions run in a background worker/runspace, keep the UI responsive, support cancel, and marshal progress/log updates through the dispatcher.
  Complexity: L

- [ ] P1 - Add protected folders and richer shortcut validation
  Why: Competitors preserve system folders and validate more than `.exe` targets; current broken-link detection misses folders, documents, app references, and URL shortcuts.
  Evidence: `StartMenuOrganizerPro.ps1:984`, ClearWinStart preserved folders, Microsoft Shell Links docs.
  Touches: `Test-ShortcutBroken`, `Get-ShortcutTarget`, folder flatten/delete logic, default settings.
  Acceptance: Built-in preserved folders are never flattened/deleted by default, `.lnk` targets are validated by target type, `.url` and app-reference shortcuts are classified, and skipped protected items are logged.
  Complexity: M

- [ ] P1 - Persist settings automatically
  Why: The UI supports import/export but does not load `$Config.ConfigFile` at startup or save edits as user preferences.
  Evidence: `StartMenuOrganizerPro.ps1:33`, `StartMenuOrganizerPro.ps1:2070`, `StartMenuOrganizerPro.ps1:2091`, Start Menu Helper options model.
  Touches: initialization, settings tab handlers, config schema/versioning.
  Acceptance: Junk patterns/categories/scope defaults persist to `%LOCALAPPDATA%`, load on startup, validate schema version, and recover cleanly from invalid JSON.
  Complexity: M

- [ ] P1 - Add Pester and PSScriptAnalyzer test gates
  Why: The repo has no tests for destructive operations and PSScriptAnalyzer already reports maintainability findings.
  Evidence: local parser pass, PSScriptAnalyzer 1.25.0 findings, Pester project.
  Touches: new tests, test fixtures, README test instructions, analyzer settings.
  Acceptance: Local test command creates disposable Start Menu fixtures, covers delete/undo/restore/flatten/rename/config paths, and analyzer warnings are either fixed or intentionally suppressed with comments.
  Complexity: L

- [ ] P2 - Package an installable release artifact
  Why: Direct peers offer installers or executable builds while this repo only documents raw `.ps1` execution.
  Evidence: Start Menu Helper setup/Inno workflow, ExplorerPatcher architecture-specific setup, README installation section.
  Touches: packaging scripts, README release instructions, artifact output.
  Acceptance: A local build creates a versioned installable artifact with Start Menu shortcut, uninstall path, execution-policy guidance, and documented admin behavior.
  Complexity: L

- [ ] P2 - Add structured file logging and crash logs
  Why: The in-window log is useful during a session but does not persist diagnostics for failed filesystem changes.
  Evidence: `StartMenuOrganizerPro.ps1:929`, ClearWinStart rotating log model.
  Touches: `Write-Log`, error handlers, `%LOCALAPPDATA%\StartMenuOrganizerPro\Logs`.
  Acceptance: Logs write to a dated file with rotation, include operation IDs from the journal, and capture unhandled exceptions before showing the user-facing error.
  Complexity: M

- [ ] P2 - Improve accessibility and localization readiness
  Why: Current WPF text is hard-coded English with custom dark controls and no visible automation labels, while mature Start Menu tools support language assets.
  Evidence: XAML resources and labels in `StartMenuOrganizerPro.ps1:104` through `StartMenuOrganizerPro.ps1:890`, Open-Shell language DLL workflow.
  Touches: XAML resources, strings, focus order, automation properties, README.
  Acceptance: User-facing strings are centralized, controls have accessible names/tooltips, keyboard focus order is predictable, contrast is verified, and future translation files can be added without editing layout logic.
  Complexity: M

- [ ] P3 - Add profile/default-user targeting after core safety lands
  Why: Admins may want to prepare another profile or the default profile, but this must wait until transaction rollback is reliable.
  Evidence: Win11Debloat advanced profile targeting, current user/system-only scope in `Get-StartMenuPaths`.
  Touches: path resolution, scope selector, elevation checks, restore/journal model.
  Acceptance: The app can scan and modify a selected offline/local profile or default profile with explicit path validation, separate backups, and clear admin gating.
  Complexity: XL
