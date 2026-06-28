# Changelog

## v0.2.0 - 2026-06-28

- Added a persistent JSON undo journal with durable undo backups under `%LOCALAPPDATA%\StartMenuOrganizerPro`.
- Journaled successful and skipped/failed destructive delete, move, rename, flatten, restore, and cleanup operations with original path, new path, backup path, action type, timestamp, and result fields.
- Added startup journal loading so the last reversible action can be undone after restarting the app.
- Added a persistent undo regression test that deletes a temp shortcut, reloads the journal, and restores it through undo.

## v0.1.0 - 2026-06-28

- Added fail-closed restore planning that stages backup contents, validates restore sources, and preserves pre-restore rollback snapshots before live Start Menu entries are replaced.
- Added restore rollback-on-failure handling and warnings for skipped or partially failed restore scopes.
- Added a local restore safety regression test and README test instructions.
