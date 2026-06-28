# Changelog

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
