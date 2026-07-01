# Start Menu Organizer

A Windows Start Menu management tool. Clean up junk, detect broken shortcuts, remove duplicates, organize by category, and take full control of your Start Menu.

![Version](https://img.shields.io/badge/Version-v0.14.0-blue)
![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

### Detection & Analysis
- **Broken Shortcut Detection** - Identifies shortcuts pointing to files/folders that no longer exist
- **Richer Shortcut Validation** - Classifies `.lnk`, `.url`, and `.appref-ms` entries, including file URL checks
- **Full Shortcut Metadata** - Extracts target, arguments, working directory, description, hotkey, and icon for `.lnk` and `.url` files
- **Risk Classification** - Flags suspicious shortcuts: script hosts, network targets, encoded commands, hidden execution, long arguments, script file targets
- **Package Provenance** - Detects winget, Chocolatey, Scoop, MSIX, and traditional installer origins
- **Duplicate Detection** - Finds multiple shortcuts targeting the same executable
- **Junk Detection** - Flags uninstall links, readme files, help docs, license files, and other clutter
- **Reparse Point Safety** - Blocks junction/symlink traversal during scans and mutations
- **Target Path Display** - See exactly where each shortcut points

### Cleanup Actions
| Action | Description |
|--------|-------------|
| Delete Selected | Remove checked items |
| Remove All Junk | Delete all detected junk items in one click |
| Remove Broken Shortcuts | Clean up shortcuts with invalid targets |
| Remove Duplicates | Keep one copy, delete redundant shortcuts |
| Flatten Single-Item Folders | Move lone shortcuts up and remove the folder |
| Remove Empty Folders | Clean up orphaned directories |
| Move All to Root | Flatten entire Start Menu by moving all shortcuts to root |

### Organization
- **11 Built-in Categories**: Development, Browsers, Communication, Media, Graphics, Office, Utilities, Gaming, System, Security, Networking
- **Move to Category** - Manually organize selected items
- **Auto-Organize All** - Automatically sort recognized apps into category folders based on pattern matching

### Batch Rename
- **Strip Version Numbers** - Remove "v1.2.3", "2024", "(1.0.0)" from names
- **Clean Up Names** - Remove "x64", "64-bit", "Microsoft", extra spaces, "- Shortcut" suffix
- **Find & Replace** - Custom text replacement in shortcut names

### User Interface
- **DataGrid View** with sortable columns: Name, Type, Status, Location, Target
- **Background Scanning** - Start Menu scans run in a cancelable worker so the window stays responsive
- **Cancelable Bulk Execution** - Reviewed operation plans run incrementally with progress and Cancel support
- **Search/Filter** (Ctrl+F) - Filter by name, path, or target in real-time
- **Profile Targeting** - Scan a selected local/offline profile or the Windows Default user profile
- **Filter Toggles** - Show/hide Shortcuts, Folders, Junk, Broken, Duplicates
- **Sort Options** - Sort by Name, Type, Status, Location, or Target
- **Preview Mode** - See what actions would do without executing them
- **Progress Bar** - Visual feedback for long operations
- **Activity Log** - Color-coded operation history
- **Accessibility Metadata** - Named controls, help text, predictable tab order, and verified contrast for core dark-theme text

### Safety Features
- **Persistent Undo Support** (Ctrl+Z) - Recover reversible deletes, moves, renames, and restores from the saved operation journal
- **Atomic File Writes** - Config and journal files use temp-file-then-rename with `.bak` recovery for crash safety
- **Backup/Restore** - Create timestamped backups before making changes
- **Fail-Closed Restore** - Validates backup contents in staging and preserves a pre-restore rollback snapshot before replacing live entries
- **Editable Transaction Plans** - Preview bulk actions into JSON plans that can be reviewed, exported, imported, and executed exactly
- **Protected Folders** - Preserves built-in folders like Startup, Windows Tools, and Administrative Tools by default
- **Persistent Settings** - Automatically saves scope, patterns, categories, and protected-folder preferences
- **Structured Logs** - Writes dated JSONL logs and crash reports under the local app data folder
- **Preview Mode** - Dry-run any operation first
- **Confirmation Dialogs** - No destructive action without explicit approval

### Export & Reporting
- **Export Scan Report** - Export current scan results as CSV or JSON
- **Enterprise Handoff** - Export cleaned shortcut inventory with deployment guidance for Intune/Start Layout
- **Virtual Group Preview** - Preview category organization without moving files
- **Rule Preset Import/Export** - Save and load junk/category/protected-folder configurations

### Elevation
- **Capability Status** - Shows which scopes are writable vs read-only in the current session
- **Relaunch as Admin** - One-click UAC elevation for full system access

### Customization (Settings Tab)
- Add/remove junk detection patterns
- Edit category patterns for auto-organization
- Export/Import configuration as JSON
- Reset to defaults

## Installation

### Option 1: Installable Package
1. Download `StartMenuOrganizer-v0.14.0.zip` from the release artifacts.
2. Extract the zip to a local folder.
3. Install for the current user:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Install-StartMenuOrganizer.ps1
```

The installer copies the app to `%LOCALAPPDATA%\Programs\Start Menu Organizer`, creates a Start Menu shortcut, and installs an uninstaller beside the app.

Uninstall:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File "$env:LOCALAPPDATA\Programs\Start Menu Organizer\Uninstall-StartMenuOrganizer.ps1"
```

### Option 2: Direct Script
1. Download `StartMenuOrganizerPro.ps1`
2. Run it from PowerShell:

```powershell
.\StartMenuOrganizerPro.ps1
```

### Verifying the Download
After downloading, verify the SHA256 hash against `SHA256SUMS.txt` inside the zip or the `.sha256` file:

```powershell
(Get-FileHash .\StartMenuOrganizer-v0.14.0.zip -Algorithm SHA256).Hash
```

### Running as Administrator
The package installer is per-user and does not require elevation. For full access to both User and System Start Menus, launch Start Menu Organizer as Administrator:

1. Right-click PowerShell → "Run as Administrator"
2. Navigate to script location and run it

## Requirements

- Windows 10 or Windows 11
- PowerShell 5.1 or higher (included with Windows)
- No additional modules or dependencies

## Usage

### Quick Start
1. Launch the script (as Admin for full access)
2. The tool scans both User and System Start Menus by default
3. Use the filter checkboxes to focus on problem items (Junk, Broken, Duplicates)
4. Select items and apply actions, or use bulk operations

### Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Del` | Delete selected items |
| `Ctrl+A` | Select all visible items |
| `Ctrl+Z` | Undo last action |
| `Ctrl+F` | Focus search box |
| `F5` | Refresh item list |

### Right-Click Context Menu
- Delete / Rename
- Open File Location
- Open Target Location
- Move to Category (submenu with all categories)
- Select All / Select None

### Recommended Workflow

1. **Create a Backup** - Click "Backup" before making changes
2. **Enable Preview Mode** - Check "Preview Mode" to test operations safely
3. **Remove Broken First** - Click "Remove Broken Shortcuts" to clean invalid links
4. **Remove Junk** - Click "Remove All Junk" to clear clutter
5. **Handle Duplicates** - Click "Remove Duplicates" to deduplicate
6. **Flatten or Organize** - Choose your preferred structure:
   - "Move All to Root" for a flat structure
   - "Auto-Organize All" for category folders
7. **Clean Names** - Use batch rename to polish shortcut names

## Start Menu Locations

The tool manages shortcuts in these directories:

| Scope | Path |
|-------|------|
| User | `%APPDATA%\Microsoft\Windows\Start Menu\Programs` |
| System | `%ProgramData%\Microsoft\Windows\Start Menu\Programs` |
| Selected Profile | `<profile root>\AppData\Roaming\Microsoft\Windows\Start Menu\Programs` |
| Default User | `%SystemDrive%\Users\Default\AppData\Roaming\Microsoft\Windows\Start Menu\Programs` |

**Note**: System, Selected Profile, and Default User modifications require Administrator privileges. Standard users can scan readable profile targets, but destructive changes, restore, and backup operations for those scopes are blocked or skipped until the app is elevated.

## Configuration

### Junk Patterns
Default patterns that flag items as junk:
```
*uninstall*, *readme*, *help*, *documentation*, *manual*,
*license*, *website*, *support*, *visit *, *about *,
*release notes*, *changelog*, *what's new*, *getting started*,
*user guide*, *online *, *web link*, *url*, *register*,
*feedback*, *update*, *check for update*
```

### Category Patterns
Each category has wildcard patterns for auto-organization. Examples:

- **Development**: `Visual Studio*`, `VS Code*`, `Git*`, `Python*`, `Docker*`
- **Browsers**: `Google Chrome*`, `Firefox*`, `Edge*`, `Brave*`
- **Gaming**: `Steam*`, `Epic Games*`, `GOG*`, `Xbox*`

### Export/Import Config
Save your customized patterns:
1. Go to Settings tab
2. Click "Export Config"
3. Save the JSON file

### Localization Overrides
UI strings are centralized in the app and can be overridden without editing layout logic. Add JSON files under `%LOCALAPPDATA%\StartMenuOrganizerPro\Localization`:

- `strings.<culture>.json`, such as `strings.en-US.json`
- `strings.<language>.json`, such as `strings.en.json`
- `strings.json` as the fallback override file

Each file maps string keys to replacement text, for example:

```json
{
  "btnBackup.Content": "Create Backup",
  "txtSearch.AutomationName": "Search Start Menu entries"
}
```

### Profile Targeting
To work against another local or offline Windows profile:

1. Open Settings.
2. Set **Profile Target** to the profile root, such as `D:\MountedUsers\Alice` or `C:\Users\Alice`.
3. Select **Selected Profile** in the Scope picker.
4. Run Start Menu Organizer as Administrator before applying changes, backup, or restore for that profile.

Use **Default User** in the Scope picker to prepare `%SystemDrive%\Users\Default`; it is also admin-gated and uses separate `DefaultUser` backup folders.

## Data Storage

| Item | Location |
|------|----------|
| Backups | `%LOCALAPPDATA%\StartMenuOrganizerPro\Backups` |
| Restore rollback snapshots | `%LOCALAPPDATA%\StartMenuOrganizerPro\Backups\_restore_rollback` |
| Undo journal | `%LOCALAPPDATA%\StartMenuOrganizerPro\undo.json` |
| Undo backups | `%LOCALAPPDATA%\StartMenuOrganizerPro\UndoBackups` |
| Config | `%LOCALAPPDATA%\StartMenuOrganizerPro\config.json` |
| Logs and crash reports | `%LOCALAPPDATA%\StartMenuOrganizerPro\Logs` |
| Localization overrides | `%LOCALAPPDATA%\StartMenuOrganizerPro\Localization` |

## Development

Run the full local gate:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1
```

Build the release package:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\Build-Release.ps1
```

Run the local restore safety regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\RestoreSafety.Tests.ps1
```

Run the persistent undo journal regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\UndoJournal.Tests.ps1
```

Run the guarded file-operation regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\GuardedFileOperation.Tests.ps1
```

Run the editable operation-plan regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\OperationPlan.Tests.ps1
```

Run the async scan wiring regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\AsyncScan.Tests.ps1
```

Run the shortcut validation regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\ShortcutValidation.Tests.ps1
```

Run the settings persistence regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\SettingsPersistence.Tests.ps1
```

Run the structured logging regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\Logging.Tests.ps1
```

Run the accessibility and localization regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\AccessibilityLocalization.Tests.ps1
```

Run the profile/default-user targeting regression test:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\ProfileTarget.Tests.ps1
```

## Troubleshooting

### "Access Denied" errors
- Run the script as Administrator for System Start Menu access

### Items not appearing
- Click "Refresh" or press F5
- Check filter toggles - some types may be hidden

### Undo not working
- Undo restores the latest reversible operation from the saved journal
- For older changes, use "Restore" to recover from a backup

### Script won't run
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Crash or failed operation diagnostics
- Check `%LOCALAPPDATA%\StartMenuOrganizerPro\Logs` for dated `.jsonl` logs and `StartMenuOrganizer-Crash-*.log` files

## Related Tools

> **Start Menu Organizer** cleans up and reorganizes what's already in your Start Menu — removing junk, broken links, and duplicates, and sorting apps into folders.
> If you want to *save, deploy, or restore* a Start Menu layout across machines or after Windows updates, see the companion tool:

**[Start Menu Manager](https://github.com/SysAdminDoc/Start-Menu-Manager)** — Export/import Start Menu layouts and taskbar configurations, manage named profiles, and perform full backups. Use Start Menu Organizer to clean up first, then Start Menu Manager to preserve and deploy the result.

## Contributing

Contributions are welcome! Feel free to submit issues and pull requests.

## License

MIT License — see the [LICENSE](LICENSE) file for details.

---

**Tip**: Always create a backup before bulk operations. While undo is available, backups provide the safest recovery option.
