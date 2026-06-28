# Roadmap

## Research-Driven Additions

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
