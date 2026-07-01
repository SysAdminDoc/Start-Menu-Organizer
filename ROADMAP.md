# Roadmap

## Research-Driven Additions

### P0

- [ ] P0 — Block reparse-point traversal before recursive scans and mutations
  Why: Root containment checks validate string paths, but recursive scan/copy/delete/restore operations can still encounter junctions, symlinks, or other reparse points.
  Evidence: `StartMenuOrganizerPro.ps1:2105`, `StartMenuOrganizerPro.ps1:2225`, `StartMenuOrganizerPro.ps1:2883`, `StartMenuOrganizerPro.ps1:4060`, `StartMenuOrganizerPro.ps1:4107`, `StartMenuOrganizerPro.ps1:4113`; PowerShell issue 621.
  Touches: `StartMenuOrganizerPro.ps1`, `tests\SafetyGuards.Tests.ps1`, `tests\Run-Tests.ps1`.
  Acceptance: Scanner reports or skips reparse points; guarded delete/move/restore refuses traversal outside approved roots; a regression fixture with a junction/symlink proves an outside-root target is untouched.
  Complexity: M

- [ ] P0 — Add full shortcut metadata extraction and risk classification
  Why: Current `.lnk` and `.url` handling only inspects target/URL, while real shortcut behavior and recent CVEs depend on arguments, working directory, icon, property store, and AppUserModelID metadata.
  Evidence: `StartMenuOrganizerPro.ps1:1956`, `StartMenuOrganizerPro.ps1:1984`, `StartMenuOrganizerPro.ps1:2903`; Microsoft Shell Link docs; MS-SHLLINK; AppUserModelID docs; CVE-2025-9491; CVE-2025-33053; Trend Micro LNK research.
  Touches: `StartMenuOrganizerPro.ps1`, scan result model, DataGrid columns, plan JSON, JSONL logs, `tests\ShortcutMetadata.Tests.ps1`.
  Acceptance: Inventory records target, arguments, working directory, icon, description, hotkey, and AppUserModelID when available; suspicious network targets, long/hidden command arguments, script host invocations, and risky `.url` files are flagged without auto-deletion; tests cover synthetic `.lnk` and `.url` fixtures.
  Complexity: L

### P1

- [ ] P1 — Resolve Start Menu roots through Known Folder APIs with safe fallbacks
  Why: Hard-coded `%PROGRAMDATA%` and profile suffix paths can miss localized, redirected, or token-specific known folders.
  Evidence: `StartMenuOrganizerPro.ps1:32`, `StartMenuOrganizerPro.ps1:34`, `StartMenuOrganizerPro.ps1:1810`; Microsoft KnownFolderID; SHGetKnownFolderPath.
  Touches: `StartMenuOrganizerPro.ps1`, profile targeting helpers, settings model, path-validation tests.
  Acceptance: User/common/profile/default-user Start Menu paths resolve through Known Folder APIs where possible, fall back to current paths when needed, and have tests for redirected or alternate profile roots.
  Complexity: M

- [ ] P1 — Make config, journal, and log writes atomic with last-known-good recovery
  Why: Direct whole-file writes can leave corrupted JSON after a crash or power loss even though invalid JSON handling exists.
  Evidence: `StartMenuOrganizerPro.ps1:1273`, `StartMenuOrganizerPro.ps1:2276`, `StartMenuOrganizerPro.ps1:2562`, `StartMenuOrganizerPro.ps1:4379`.
  Touches: `StartMenuOrganizerPro.ps1`, config persistence tests, undo journal tests, logging tests.
  Acceptance: Writes use temp files plus replace/rename and `.bak` preservation; startup recovers from a partial config/journal file using the last good copy; tests simulate truncated JSON.
  Complexity: S

- [ ] P1 — Emit release integrity artifacts and optional Authenticode signing
  Why: The packaged shortcut runs with `-ExecutionPolicy Bypass`, and the release package has no hash manifest or signing status for users on RemoteSigned/AllSigned policies.
  Evidence: `tools\Build-Release.ps1:67`; PowerShell about_Signing; PowerShell about_Execution_Policies.
  Touches: `tools\Build-Release.ps1`, `README.md`, `CHANGELOG.md`, release package tests.
  Acceptance: Release builds emit a SHA256 manifest, sign scripts when a configured certificate is available, report unsigned status when not available, and README shows local verification commands.
  Complexity: M

- [ ] P1 — Add human-readable scan and operation audit report export
  Why: JSON plans, journals, and JSONL logs exist, but users need a before/after report they can review, archive, or hand to another admin without parsing internal files.
  Evidence: README workflows; `StartMenuOrganizerPro.ps1` plan/journal/logging code; Start11 backup/export features; Reddit Start menu backup thread.
  Touches: `StartMenuOrganizerPro.ps1`, report exporter, UI actions tab, tests for HTML/CSV/JSON report output.
  Acceptance: Current scan or completed operation exports a report with item path, action, target metadata, warnings, skipped/protected items, result, timestamp, and source scope; tests verify schema and output redacts or clearly labels user-profile roots.
  Complexity: M

### P2

- [ ] P2 — Track package-manager provenance and recreated shortcut policy
  Why: Chocolatey and Scoop users report package updates recreating Start Menu shortcuts after users move or delete them; the app needs policy-aware handling instead of treating every reappearance as fresh clutter.
  Evidence: Chocolatey issue 2016; Scoop issue 4562; current scan groups by path/target only.
  Touches: `StartMenuOrganizerPro.ps1`, config schema, journal schema, plan generation tests.
  Acceptance: Scanner records likely winget/choco/scoop/vendor provenance and last-seen/original locations; rules can ignore, restore, quarantine, or re-plan recreated shortcuts without duplicating old actions.
  Complexity: L

- [ ] P2 — Replace hard-coded cleanup tables with an ordered rule engine and preset import/export
  Why: Category, junk, and protected-folder data exists but rule precedence and conflicts are not inspectable or portable.
  Evidence: `StartMenuOrganizerPro.ps1:58`, `StartMenuOrganizerPro.ps1:72`, `StartMenuOrganizerPro.ps1:84`; Start Menu Helper options; Win11Debloat presets.
  Touches: `StartMenuOrganizerPro.ps1`, settings tab, config schema, rule tests, README.
  Acceptance: Rules have explicit order, category/protected/junk precedence, conflict diagnostics in preview, export/importable presets, and tests for overlapping patterns.
  Complexity: M

- [ ] P2 — Convert destructive modal confirmations into plan-backed status flow
  Why: The app already has editable operation plans and logs, but destructive actions still rely on modal confirmation dialogs that interrupt review and execution.
  Evidence: README confirmation workflow; `StartMenuOrganizerPro.ps1` action handlers using `MessageBox.Show`; existing plan preview/execution model.
  Touches: `StartMenuOrganizerPro.ps1`, action tab XAML, status/log panel, operation plan tests.
  Acceptance: Destructive buttons create or append reviewed operation plans by default; execution reports in-window status/log feedback; ad hoc modal confirmations are removed for plan-backed operations.
  Complexity: L

- [ ] P2 — Add localization template export and completeness validation
  Why: Centralized strings and accessibility names exist, but override files have no template export or missing/orphaned-key validator.
  Evidence: `StartMenuOrganizerPro.ps1` `DefaultUiStrings`; localization override path in config; WPF accessibility docs requiring accurate localized UIA names.
  Touches: `StartMenuOrganizerPro.ps1`, `tests\AccessibilityLocalization.Tests.ps1`, `README.md`.
  Acceptance: A local command or test exports `strings.template.json`, validates missing and orphaned localization keys including UIA strings, and documents the override workflow.
  Complexity: S

- [ ] P2 — Add runtime WPF smoke and accessibility regression checks
  Why: Current tests parse and lint code, but they do not render tabs/themes or verify focusable controls, clipping, visibility, and UI Automation names at runtime.
  Evidence: `tests\Run-Tests.ps1`; `tests\AccessibilityLocalization.Tests.ps1`; WPF accessibility docs.
  Touches: `tests\Run-Tests.ps1`, new WPF smoke tests, `StartMenuOrganizerPro.ps1` test-mode launch hook if needed.
  Acceptance: Local tests open the WPF shell in test mode, walk each major tab, assert key controls are visible and named, and fail on missing UIA labels or obvious layout clipping.
  Complexity: M

- [ ] P2 — Add Pester code coverage reporting for core safety functions
  Why: Parser, Pester, and ScriptAnalyzer checks run now, but there is no coverage signal for safety-critical branches.
  Evidence: `tests\Run-Tests.ps1`; Pester code coverage docs; safety helpers in `StartMenuOrganizerPro.ps1`.
  Touches: `tests\Run-Tests.ps1`, Pester configuration, README test instructions.
  Acceptance: Local test run emits coverage for core functions, documents exclusions for WPF-only code, and enforces an initial threshold that does not require GitHub Actions.
  Complexity: S

- [ ] P2 — Add elevated relaunch and capability status panel
  Why: README instructs users to manually run as Administrator for system-wide changes, but the UI can make current capability and elevation path explicit.
  Evidence: README administrator guidance; `StartMenuOrganizerPro.ps1` admin checks and `Test-MutationRootRequiresAdministrator`.
  Touches: `StartMenuOrganizerPro.ps1`, settings/actions UI, admin-scope tests.
  Acceptance: Non-admin sessions show which scopes are read-only or mutation-disabled and provide a relaunch-elevated action using `ShellExecute` `runas`; tests verify command construction and disabled-state behavior.
  Complexity: S

- [ ] P2 — Add enterprise handoff export for cleaned Start layouts
  Why: Microsoft supports Start layout deployment, and commercial tools sell role/kiosk layout export; this app can hand off cleaned filesystem shortcuts without becoming a shell replacement.
  Evidence: Microsoft Start layout deployment docs; Start11 deployment/export features; README companion Start Menu Manager note.
  Touches: `StartMenuOrganizerPro.ps1`, report/export code, README.
  Acceptance: After cleanup, the app exports a handoff bundle with cleaned shortcut inventory, warnings, and Start layout/Intune guidance or companion import payload; no direct Explorer hooking is introduced.
  Complexity: L

### P3

- [ ] P3 — Add non-mutating virtual grouping preview
  Why: Start Menu X shows virtual groups avoid installer/uninstaller folder drift, but this project should use that idea only as preview/reporting rather than replacing real cleanup.
  Evidence: Start Menu X virtual groups documentation; current operation-plan preview model.
  Touches: `StartMenuOrganizerPro.ps1`, preview model, report exporter.
  Acceptance: Users can preview category/group organization without moving files and export the preview as a report or rule preset; UI clearly distinguishes virtual preview from real filesystem operations.
  Complexity: L

- [ ] P3 — Prepare package metadata for winget, Chocolatey, and Scoop after signed releases stabilize
  Why: Comparable Windows utilities support multiple install channels, but package-manager distribution should wait for reliable hashes and signing.
  Evidence: PowerToys install channels; Chocolatey and Scoop shortcut recreation issues; release integrity roadmap item.
  Touches: `tools\Build-Release.ps1`, package metadata files, README install section.
  Acceptance: Package metadata can be generated locally from the signed release artifact and SHA256 manifest, with install/uninstall behavior documented and tested.
  Complexity: M
