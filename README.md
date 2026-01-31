# Start Menu Organizer

A Windows Start Menu management tool. Clean up junk, detect broken shortcuts, remove duplicates, organize by category, and take full control of your Start Menu.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue)
![Windows](https://img.shields.io/badge/Windows-10%20%7C%2011-0078D6)
![License](https://img.shields.io/badge/License-MIT-green)

## Features

### Detection & Analysis
- **Broken Shortcut Detection** - Identifies shortcuts pointing to files/folders that no longer exist
- **Duplicate Detection** - Finds multiple shortcuts targeting the same executable
- **Junk Detection** - Flags uninstall links, readme files, help docs, license files, and other clutter
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
- Modern dark theme with GitHub-inspired colors
- **DataGrid View** with sortable columns: Name, Type, Status, Location, Target
- **Search/Filter** (Ctrl+F) - Filter by name, path, or target in real-time
- **Filter Toggles** - Show/hide Shortcuts, Folders, Junk, Broken, Duplicates
- **Sort Options** - Sort by Name, Type, Status, Location, or Target
- **Preview Mode** - See what actions would do without executing them
- **Progress Bar** - Visual feedback for long operations
- **Activity Log** - Color-coded operation history

### Safety Features
- **Undo Support** (Ctrl+Z) - Recover from accidental deletions
- **Backup/Restore** - Create timestamped backups before making changes
- **Preview Mode** - Dry-run any operation first
- **Confirmation Dialogs** - No destructive action without explicit approval

### Customization (Settings Tab)
- Add/remove junk detection patterns
- Edit category patterns for auto-organization
- Export/Import configuration as JSON
- Reset to defaults

## Installation

No installation required. Simply download and run.

### Option 1: Direct Download
1. Download `StartMenuOrganizerPro.ps1`
2. Right-click and select "Run with PowerShell"

### Option 2: From PowerShell
```powershell
# Navigate to the script location
cd "C:\Path\To\Script"

# Run the script
.\StartMenuOrganizerPro.ps1
```

### Running as Administrator
For full access to both User and System Start Menus, run as Administrator:

1. Right-click PowerShell → "Run as Administrator"
2. Navigate to script location and run it

Or right-click the script → "Run with PowerShell as Administrator"

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

**Note**: System Start Menu modifications require Administrator privileges.

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

Import on another machine:
1. Go to Settings tab
2. Click "Import Config"
3. Select your JSON file

## Data Storage

| Item | Location |
|------|----------|
| Backups | `%LOCALAPPDATA%\StartMenuOrganizerPro\Backups` |
| Config | `%LOCALAPPDATA%\StartMenuOrganizerPro\config.json` |

## Troubleshooting

### "Access Denied" errors
- Run the script as Administrator for System Start Menu access

### Items not appearing
- Click "Refresh" or press F5
- Check filter toggles - some types may be hidden

### Undo not working
- Undo only works for deletions made in the current session
- For older changes, use "Restore" to recover from a backup

### Script won't run
PowerShell execution policy may be blocking scripts:
```powershell
# Check current policy
Get-ExecutionPolicy

# Allow scripts for current user (if needed)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## Contributing

Contributions are welcome! Feel free to submit issues and pull requests.

## License

MIT License - Feel free to use, modify, and distribute.

## Author

Matt | Maven Imaging

---

**Tip**: Always create a backup before bulk operations. While undo is available, backups provide the safest recovery option.
