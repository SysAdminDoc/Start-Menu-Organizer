# Changelog

## v0.1.0 - 2026-06-28

- Added fail-closed restore planning that stages backup contents, validates restore sources, and preserves pre-restore rollback snapshots before live Start Menu entries are replaced.
- Added restore rollback-on-failure handling and warnings for skipped or partially failed restore scopes.
- Added a local restore safety regression test and README test instructions.
