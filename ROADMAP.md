# Roadmap

## Research-Driven Additions

- [ ] P3 - Add profile/default-user targeting after core safety lands
  Why: Admins may want to prepare another profile or the default profile, but this must wait until transaction rollback is reliable.
  Evidence: Win11Debloat advanced profile targeting, current user/system-only scope in `Get-StartMenuPaths`.
  Touches: path resolution, scope selector, elevation checks, restore/journal model.
  Acceptance: The app can scan and modify a selected offline/local profile or default profile with explicit path validation, separate backups, and clear admin gating.
  Complexity: XL
