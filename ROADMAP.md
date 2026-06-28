# Roadmap

## Research-Driven Additions

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
