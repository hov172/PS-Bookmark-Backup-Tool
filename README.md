# 📑 Bookmark Backup Tool v5.2.1 
  ## PowerShell Enhanced Edition

PowerShell tool that lets you **export and import bookmarks** for **Google Chrome**, **Microsoft Edge**, and **Mozilla Firefox**.  
It supports GUI for interactive use and CLI for automation, offers scheduled backups, multi-profile handling, **HTML conversion**, **ZIP archives**, **browser auto-close**, file integrity checks, and comprehensive logging.  
It's designed to be **robust in enterprise environments** (network paths, permissions, and process safety) while staying simple for individual users.

**I built both Windows and macOS applications based on this script. These apps allow users to manage their own bookmarks, and both versions include command-line options for administrators.**

- 💻 **MacOS**: [https://github.com/hov172/MacOS-Bookmarks-Backup-Tool](https://github.com/hov172/MacOS-Bookmarks-Backup-Tool)  
- 🖥️ **Windows**: [https://github.com/hov172/Win-Bookmarks-Backup-Tool/tree/main](https://github.com/hov172/Win-Bookmarks-Backup-Tool/tree/main)

---

<img width="579" height="410" alt="bookmark" src="https://github.com/user-attachments/assets/f9c5ad00-9372-4d14-9c68-5846da5c95fe" />

---

## 📝 Changelog

**v5.2 — December 10, 2025**

- 🔄 **Browser Auto-Close**: Graceful close with 3s wait, force kill if needed
- 📄 **HTML Export/Conversion**: Real Chrome JSON → HTML and Firefox SQLite → HTML
- 📦 **ZIP Archive Support**: Create and extract ZIP archives of bookmarks
- 🌐 **All-Profiles Support**: Export/import from all browser profiles
- ⚠️ **Interactive Browser Warnings**: Y/N prompts when browsers are running
- 🔄 **Retry Logic**: Auto-retry after browser closure with 2s delay
- 📁 **ZIP Import**: Direct import from ZIP archives with auto-detection
- 🔧 **Auto-Install SQLite**: Downloads System.Data.SQLite from NuGet if missing

**v5.0 — October 25, 2025**

- Made into a PowerShell Module
- Publish installation-powershell-gallery

**v5.0 — September 22, 2025**

- Enhanced GUI  
- Network path detection with retry  
- Process safety checks  
- Auto backups before import  
- File integrity validation  
- Comprehensive logging  
- Scheduling improvements

---

## 📚 Table of Contents

- [🧭 Overview](#-overview)
- [✨ Key Features](#-key-features)
- [🆕 What's New in v5.2](#-whats-new-in-v52)
- [🧰 Requirements](#-requirements)
- [🌐 Supported Browsers & Files](#-supported-browsers--files)
- [📦 Installation (PowerShell Gallery)](#-installation-powershell-gallery)
- [🧭 Available Commands](#-available-commands)
- [⚡ Quick Start](#-quick-start)
- [🖱️ Usage](#️-usage)
  - [GUI Mode (Beginner-friendly)](#gui-mode-beginner-friendly)
  - [CLI / Silent Mode](#cli--silent-mode)
  - [🧾 Parameters](#-parameters)
  - [🧰 Command-Line Examples](#-command-line-examples)
- [📄 HTML Export Feature](#-html-export-feature)
- [📦 ZIP Archive Feature](#-zip-archive-feature)
- [🔄 Browser Auto-Close Feature](#-browser-auto-close-feature)
- [🕑 Scheduling (Automatic Backups)](#-scheduling-automatic-backups)
- [⚙️ Configuration](#️-configuration)
- [🧠 How It Works](#-how-it-works)
- [📂 File Locations](#-file-locations)
- [🧭 Troubleshooting](#-troubleshooting)
- [❓ FAQ](#-faq)
- [🔐 Security Notes](#-security-notes)
- [⚠️ Known Limitations](#️-known-limitations)
- [🧹 Uninstall / Cleanup](#-uninstall--cleanup)
- [📜 License](#-license)
- [🙌 Credits](#-credits)

---

## 🧭 Overview

**Bookmark Backup Tool v5.2** is a PowerShell module that lets you export and import bookmarks for **Google Chrome**, **Microsoft Edge**, and **Mozilla Firefox**.  
It supports **GUI** for interactive use and **CLI** for automation, offers **scheduled backups**, **multi-profile handling**, **HTML conversion**, **ZIP archives**, **browser auto-close**, **file integrity checks**, and **comprehensive logging**.  
It's built to work smoothly in both **enterprise** and **home** environments.

---

## ✨ Key Features

- **Multi-browser support** (Chrome, Edge, Firefox)
- **GUI** (Windows Forms) and **CLI** (silent mode)
- **Dual-format export**: Native format (JSON/SQLite) + HTML
- **HTML-only export mode**: Lightweight, universal bookmark format
- **ZIP archive support**: Bundle all bookmarks into compressed archives
- **Browser auto-close**: Automatically closes browsers when needed for safe imports
- **Multi-profile support**: Export/import all browser profiles or just the latest
- **Automatic path detection** with network share preference and Desktop fallback
- **Automatic backup** before imports (maintains up to 10 rolling backups)
- **File integrity verification** for JSON and SQLite files
- **Task Scheduler integration** (Daily/Weekly/Monthly)
- **Configuration file support** with sensible defaults
- **Verbose logging** with retention policy
- **Auto-install SQLite**: Downloads System.Data.SQLite from NuGet if needed
- **Safety mechanisms**: Browser process detection, retry logic, WhatIf/Force switches
- **PowerShell Module** for easy distribution and updates

---

## 🆕 What's New in v5.2

### 📄 HTML Conversion
Export bookmarks in **universal HTML format** (Netscape Bookmark format) that works across all browsers:
- **Chrome/Edge**: Converts JSON bookmarks to HTML automatically
- **Firefox**: Converts SQLite database to HTML (requires System.Data.SQLite)
- **Dual-format exports**: Get both native format AND HTML in one operation
- **HTML-only mode**: Use `-HtmlOnly` for lightweight backups

### 📦 ZIP Archive Support
- **Create archives**: Use `-CreateZip` to bundle all exported bookmarks
- **Import from ZIP**: Automatically extracts and imports bookmarks from ZIP files
- **Auto-detection**: Smart detection of browser types within archives
- **Space-efficient**: Compress multiple bookmark backups into single files

### 🔄 Browser Auto-Close
- **Graceful shutdown**: Attempts clean browser exit with 3-second wait
- **Force kill fallback**: Ensures browsers close even if unresponsive
- **Interactive prompts**: Y/N confirmation before closing browsers in GUI mode
- **Silent mode support**: Auto-close with `-CloseBrowserIfRunning` flag
- **Multi-process handling**: Closes all related browser processes (helpers, renderers, etc.)

### 🌐 All-Profiles Support
- **Export all profiles**: Use `-AllProfiles` to backup every browser profile
- **Profile naming**: Automatically names exports by profile (Default, Profile 1, etc.)
- **Latest profile detection**: Default behavior exports most recently used profile
- **Import to all profiles**: Restore bookmarks across all profiles simultaneously

### 🔧 Auto-Install SQLite
- **Zero configuration**: Automatically downloads System.Data.SQLite if missing
- **NuGet integration**: Fetches official SQLite package (v1.0.118)
- **Multi-architecture**: Supports both x86 and x64 systems
- **Fallback handling**: Gracefully degrades if SQLite unavailable

---

## 🧰 Requirements

- **Windows 10 or 11**  
- **PowerShell 5.1+**  
- **.NET Framework 4.6+** (usually pre-installed)
- **System Assemblies**: `System.Windows.Forms`, `System.Drawing`, `System.IO.Compression.FileSystem`  
- **File system access** to browser profiles and target paths  
- **Admin rights** for scheduled tasks (optional, only for task creation)
- **Internet connection** (optional, for auto-downloading System.Data.SQLite)

---

## 🌐 Supported Browsers & Files

| Browser  | Native File                   | Export Formats           | Native Location                                                                 |
|----------|-------------------------------|--------------------------|----------------------------------------------------------------------------------|
| Chrome   | `Bookmarks` (JSON)            | JSON + HTML              | `%LOCALAPPDATA%\Google\Chrome\User Data\<Profile>\Bookmarks`                    |
| Edge     | `Bookmarks` (JSON)            | JSON + HTML              | `%LOCALAPPDATA%\Microsoft\Edge\User Data\<Profile>\Bookmarks`                   |
| Firefox  | `places.sqlite` (SQLite DB)   | SQLite + HTML            | `%APPDATA%\Mozilla\Firefox\Profiles\<Profile>\places.sqlite`                    |

### Exported File Naming Convention

**Single Profile Exports:**
- `Chrome_BookmarkData_2025-12-10_14-30-00.json`
- `Chrome_Bookmarks_2025-12-10_14-30-00.html`
- `Edge_BookmarkData_2025-12-10_14-30-00.json`
- `Edge_Bookmarks_2025-12-10_14-30-00.html`
- `Firefox_BookmarkData_2025-12-10_14-30-00.sqlite`
- `Firefox_Bookmarks_2025-12-10_14-30-00.html`

**Multi-Profile Exports (with `-AllProfiles`):**
- `Chrome-Default_Profile_BookmarkData_2025-12-10_14-30-00.json`
- `Chrome-Profile_1_Bookmarks_2025-12-10_14-30-00.html`

**ZIP Archive:**
- `BookmarkBackup_2025-12-10_14-30-00.zip`

---

## 📦 Installation (PowerShell Gallery)

👉 [PowerShell Gallery Package](https://www.powershellgallery.com/packages/BookmarkBackupTool/5.2.0)

```powershell
# Install for current user
Install-Module -Name BookmarkBackupTool -Scope CurrentUser -Force

# Or install system-wide (Admin required)
Install-Module -Name BookmarkBackupTool -Scope AllUsers -Force
```

> 💡 First-time users may need to trust PSGallery:
> ```powershell
> Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
> ```

**Update the module:**
```powershell
Update-Module -Name BookmarkBackupTool
```

**Uninstall the module:**
```powershell
Uninstall-Module -Name BookmarkBackupTool
```

---

## 🧭 Available Commands

| Command                         | Description                                                           |
|----------------------------------|-----------------------------------------------------------------------|
| `Export-Bookmarks`              | Exports bookmarks from supported browsers to backup files (JSON/SQLite + HTML). |
| `Import-Bookmarks`              | Imports bookmarks from backup files or ZIP archives.                  |
| `Get-HomeSharePath`            | Returns the user's home or shared folder path with network fallback.  |
| `Test-BrowserInstalled`        | Checks if a supported browser is installed on the system.             |
| `Test-BrowserRunning`          | Checks if a browser is currently running (multi-process detection).   |
| `Close-Browser`                | Gracefully closes a browser with force-kill fallback.                 |
| `Get-BrowserProfiles`          | Lists all available browser profiles with last-used timestamps.       |
| `Get-AllBrowserProfiles`       | Gets detailed profile information for multi-profile operations.       |
| `Backup-ExistingBookmarks`     | Creates a timestamped snapshot of current bookmarks (maintains 10 rolling backups). |
| `ConvertTo-ChromeHtml`         | Converts Chrome JSON bookmarks to Netscape HTML format.               |
| `ConvertTo-FirefoxHtml`        | Converts Firefox SQLite database to Netscape HTML format.             |
| `New-ZipArchive`               | Creates a ZIP archive from exported bookmark files.                   |
| `Expand-ZipArchive`            | Extracts bookmark files from ZIP archives.                            |
| `Import-FromZip`               | Imports bookmarks directly from ZIP archives with auto-detection.     |
| `Show-BookmarkGUI`             | Launches the graphical user interface (v5.2 with new checkboxes).    |
| `New-BookmarkScheduledTask`    | Creates a Windows scheduled task for automated backups.               |
| `Remove-BookmarkScheduledTask` | Removes the scheduled backup task from Task Scheduler.                |
| `Get-BookmarkConfiguration`    | Displays current configuration settings.                              |
| `Set-BookmarkConfiguration`    | Updates configuration file with new settings.                         |
| `Test-BookmarkPrerequisites`   | Checks environment readiness (PowerShell version, .NET, permissions). |
| `Test-BookmarkFileIntegrity`   | Validates bookmark file structure (JSON/SQLite format checks).        |
| `Install-SQLiteIfMissing`      | Auto-downloads and installs System.Data.SQLite from NuGet.            |

To list all commands:

```powershell
Get-Command -Module BookmarkBackupTool
```

To get detailed help for any command:

```powershell
Get-Help Export-Bookmarks -Full
Get-Help Import-Bookmarks -Examples
```

---

## ⚡ Quick Start

### Launch GUI (Easiest Method)
```powershell
Show-BookmarkGUI
```

### Quick Export (All Browsers with HTML)
```powershell
Export-Bookmarks -Chrome -Edge -Firefox -TargetPath "$env:USERPROFILE\Desktop\BookmarkBackup"
```

### Quick Export (HTML Only - Lightweight)
```powershell
Export-Bookmarks -Chrome -Edge -Firefox -HtmlOnly -TargetPath "$env:USERPROFILE\Desktop\BookmarkBackup"
```

### Quick Export with ZIP Archive
```powershell
Export-Bookmarks -Chrome -Edge -Firefox -CreateZip -TargetPath "$env:USERPROFILE\Desktop\BookmarkBackup"
```

### Quick Import (Auto-Close Browsers)
```powershell
Import-Bookmarks -Chrome -Edge -Firefox -CloseBrowserIfRunning -TargetPath "$env:USERPROFILE\Desktop\BookmarkBackup"
```

---

## 🖱️ Usage

### GUI Mode (Beginner-friendly)

Launch the graphical interface for point-and-click operations:

```powershell
Show-BookmarkGUI
```

**GUI Features (v5.2):**
- ✅ Browser selection checkboxes (Chrome, Edge, Firefox)
- ✅ **Process all profiles** checkbox (new in v5.2)
- ✅ **HTML-only export** checkbox (new in v5.2)
- ✅ **Create ZIP archive** checkbox (new in v5.2)
- 📁 Browse button for custom save/load paths
- 📤 Export Bookmarks button
- 📥 Import Bookmarks button
- 💡 Helpful tip text about dual-format backups
- 🔄 Auto-closes browsers with interactive Y/N prompts

### CLI / Silent Mode

For automation, scripting, and scheduled tasks:

```powershell
# Basic silent export
.\BookmarkTool.ps1 -Silent -Action Export -Chrome -Edge -TargetPath "C:\Backups"

# Basic silent import
.\BookmarkTool.ps1 -Silent -Action Import -Chrome -TargetPath "C:\Backups"
```

---

### 🧾 Parameters

| Parameter               | Type       | Description                                                       |
|-------------------------|------------|-------------------------------------------------------------------|
| `-Silent`               | Switch     | Run without GUI (required for CLI mode)                           |
| `-Action`               | String     | `Export` or `Import` (required with `-Silent`)                     |
| `-Chrome`               | Switch     | Include Google Chrome bookmarks                                   |
| `-Edge`                 | Switch     | Include Microsoft Edge bookmarks                                  |
| `-Firefox`              | Switch     | Include Mozilla Firefox bookmarks                                 |
| `-TargetPath`           | String     | Specific path for bookmark storage (overrides auto-detection)     |
| `-HtmlOnly`             | Switch     | Export only HTML format (lightweight, no native files) **NEW**    |
| `-AllProfiles`          | Switch     | Process all browser profiles instead of just the latest **NEW**   |
| `-CreateZip`            | Switch     | Create ZIP archive of exported bookmarks **NEW**                  |
| `-CloseBrowserIfRunning`| Switch     | Auto-close browsers without prompts (silent mode) **NEW**         |
| `-Force`                | Switch     | Force operations without confirmations                            |
| `-CreateScheduledTask`  | Switch     | Create scheduled task for automatic backups                       |
| `-ScheduleFrequency`    | String     | `Daily`, `Weekly`, or `Monthly` (default: Daily)                  |
| `-ConfigPath`           | String     | Path to custom configuration file                                 |
| `-WhatIf`               | Switch     | Preview operations without making changes                         |
| `-Verbose`              | Switch     | Enable detailed logging output                                    |

---

### 🧰 Command-Line Examples

#### Basic Export Operations

```powershell
# Export Chrome bookmarks to Desktop
Export-Bookmarks -Chrome -TargetPath "$env:USERPROFILE\Desktop\Backups"

# Export all browsers with default path detection
Export-Bookmarks -Chrome -Edge -Firefox

# Export with verbose logging
Export-Bookmarks -Chrome -Edge -Verbose
```

#### HTML Export Operations (NEW in v5.2)

```powershell
# Export only HTML (no native JSON/SQLite files)
Export-Bookmarks -Chrome -Edge -Firefox -HtmlOnly -TargetPath "C:\Backups"

# HTML export is ideal for:
# - Cross-browser compatibility
# - Smaller file sizes
# - Universal bookmark format
# - Archival purposes
```

#### Multi-Profile Operations (NEW in v5.2)

```powershell
# Export bookmarks from ALL Chrome profiles
Export-Bookmarks -Chrome -AllProfiles -TargetPath "C:\Backups"

# Export all profiles from all browsers
Export-Bookmarks -Chrome -Edge -Firefox -AllProfiles -TargetPath "C:\Backups"

# Import to all Edge profiles
Import-Bookmarks -Edge -AllProfiles -TargetPath "C:\Backups"
```

#### ZIP Archive Operations (NEW in v5.2)

```powershell
# Export with ZIP archive creation
Export-Bookmarks -Chrome -Edge -Firefox -CreateZip -TargetPath "C:\Backups"

# Import directly from ZIP file
Import-Bookmarks -Chrome -Edge -Firefox -TargetPath "C:\Backups\BookmarkBackup_2025-12-10_14-30-00.zip"

# The ZIP will contain all exported files:
# - Chrome_BookmarkData_*.json
# - Chrome_Bookmarks_*.html
# - Edge_BookmarkData_*.json
# - Edge_Bookmarks_*.html
# - Firefox_BookmarkData_*.sqlite
# - Firefox_Bookmarks_*.html
```

#### Import Operations with Browser Auto-Close (NEW in v5.2)

```powershell
# Import with automatic browser closure (silent mode)
Import-Bookmarks -Chrome -Edge -CloseBrowserIfRunning -TargetPath "C:\Backups"

# Force import without confirmations
Import-Bookmarks -Firefox -Force -TargetPath "C:\Backups"

# Import from specific files
Import-Bookmarks -Chrome -TargetPath "C:\Backups\Chrome-Bookmarks.json"
```

#### Combined Operations

```powershell
# Full-featured export: All browsers, all profiles, HTML, ZIP
Export-Bookmarks -Chrome -Edge -Firefox -AllProfiles -CreateZip -TargetPath "C:\Backups"

# Silent mode with all features
.\BookmarkTool.ps1 -Silent -Action Export -Chrome -Edge -Firefox -AllProfiles -CreateZip -HtmlOnly -TargetPath "C:\Backups"
```

#### Preview Operations (WhatIf)

```powershell
# Preview export without making changes
Export-Bookmarks -Chrome -Edge -WhatIf

# Preview import
Import-Bookmarks -Firefox -TargetPath "C:\Backups" -WhatIf
```

#### Network Path Examples

```powershell
# Use specific network share
Export-Bookmarks -Chrome -Edge -TargetPath "\\server\backups\bookmarks"

# Let tool auto-detect network path (uses HOMESHARE environment variable)
Export-Bookmarks -Chrome -Edge
```

---

## 📄 HTML Export Feature

### Overview
Version 5.2 introduces **dual-format export** capability, creating both native bookmark files and universal HTML format simultaneously.

### HTML Format Benefits
- ✅ **Universal compatibility**: Works with any browser
- ✅ **Human-readable**: View bookmarks in any text editor
- ✅ **Standard format**: Netscape Bookmark File Format
- ✅ **Portable**: Share across platforms and browsers
- ✅ **Smaller size**: More compact than native formats
- ✅ **Archival**: Long-term storage format

### How It Works

**Chrome/Edge (JSON → HTML):**
- Reads Chrome's JSON bookmark structure
- Parses bookmark bar and other folders
- Converts to hierarchical HTML
- Preserves folder structure and URLs

**Firefox (SQLite → HTML):**
- Queries `places.sqlite` database
- Extracts bookmarks via SQL
- Converts to HTML format
- Requires System.Data.SQLite (auto-downloads if missing)

### Export Modes

**Default Mode (Dual-Format):**
```powershell
# Creates BOTH native files AND HTML
Export-Bookmarks -Chrome -Edge -Firefox -TargetPath "C:\Backups"

# Result:
# ✓ Chrome_BookmarkData_2025-12-10.json
# ✓ Chrome_Bookmarks_2025-12-10.html
# ✓ Edge_BookmarkData_2025-12-10.json
# ✓ Edge_Bookmarks_2025-12-10.html
# ✓ Firefox_BookmarkData_2025-12-10.sqlite
# ✓ Firefox_Bookmarks_2025-12-10.html
```

**HTML-Only Mode:**
```powershell
# Creates ONLY HTML files (lightweight)
Export-Bookmarks -Chrome -Edge -Firefox -HtmlOnly -TargetPath "C:\Backups"

# Result:
# ✓ Chrome_Bookmarks_2025-12-10.html
# ✓ Edge_Bookmarks_2025-12-10.html
# ✓ Firefox_Bookmarks_2025-12-10.html
```

### HTML File Structure

```html
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>
<DL><p>
    <DT><H3>Bookmarks Bar</H3>
    <DL><p>
        <DT><A HREF="https://example.com" ADD_DATE="1701234567">Example Site</A>
        <DT><H3>Work</H3>
        <DL><p>
            <DT><A HREF="https://work.example.com">Work Site</A>
        </DL><p>
    </DL><p>
</DL><p>
```

### SQLite Auto-Install

If Firefox HTML conversion is needed but System.Data.SQLite is missing:

1. **Automatic download** from NuGet (Stub.System.Data.SQLite.Core.NetFramework v1.0.118)
2. **Architecture detection** (x86/x64)
3. **Local installation** to `lib` folder in script directory
4. **Graceful fallback** if download fails (exports SQLite only)

```powershell
# Check if SQLite is available
if ($script:SQLiteAvailable) {
    Write-Host "✓ System.Data.SQLite is available for Firefox HTML conversion"
} else {
    Write-Host "⚠ System.Data.SQLite not available - Firefox exports SQLite only"
}
```

---

## 📦 ZIP Archive Feature

### Overview
Create compressed archives of bookmark exports for easy storage, transfer, and backup management.

### Creating ZIP Archives

```powershell
# Export with automatic ZIP creation
Export-Bookmarks -Chrome -Edge -Firefox -CreateZip -TargetPath "C:\Backups"

# Result:
# ✓ All individual bookmark files created
# ✓ BookmarkBackup_2025-12-10_14-30-00.zip (containing all files)
```

### ZIP Archive Contents

The ZIP archive contains all exported files:
- Native bookmark files (JSON/SQLite)
- HTML converted files
- All profiles if `-AllProfiles` is used

Example ZIP contents:
```
BookmarkBackup_2025-12-10_14-30-00.zip
├── Chrome_BookmarkData_2025-12-10_14-30-00.json
├── Chrome_Bookmarks_2025-12-10_14-30-00.html
├── Edge_BookmarkData_2025-12-10_14-30-00.json
├── Edge_Bookmarks_2025-12-10_14-30-00.html
├── Firefox_BookmarkData_2025-12-10_14-30-00.sqlite
└── Firefox_Bookmarks_2025-12-10_14-30-00.html
```

### Importing from ZIP

```powershell
# Automatic extraction and import
Import-Bookmarks -Chrome -Edge -Firefox -TargetPath "C:\Backups\BookmarkBackup_2025-12-10.zip"

# The tool will:
# 1. Detect that path is a ZIP file
# 2. Extract to temporary folder
# 3. Auto-detect browser files
# 4. Import to respective browsers
# 5. Clean up temporary files
```

### ZIP Auto-Detection

The import function automatically detects:
- Chrome files (matching pattern: `Chrome.*\.json`)
- Edge files (matching pattern: `Edge.*\.json`)
- Firefox files (matching pattern: `Firefox.*\.sqlite`)

### Benefits of ZIP Archives

- 💾 **Space efficient**: Compressed storage
- 📦 **Single file**: Easy to transfer and manage
- 🔒 **Integrity**: All bookmarks bundled together
- 📧 **Email-friendly**: Smaller file sizes
- 💿 **Archival**: Perfect for long-term backup storage

---

## 🔄 Browser Auto-Close Feature

### Overview
Version 5.2 introduces **intelligent browser management** that safely closes browsers when needed for import operations.

### How It Works

**Detection:**
1. Checks for all browser-related processes:
   - Chrome: `chrome`, `GoogleChromeHelper`, `crashpad_handler`
   - Edge: `msedge`, `MicrosoftEdge`, `MicrosoftEdgeWebView2`, `identity_helper`
   - Firefox: `firefox`, `plugin-container`, `crashreporter`, `updater`

2. **Interactive Mode (GUI):**
   ```
   ⚠️ WARNING: Chrome, Edge is currently running!
   The browser(s) must be closed to import bookmarks safely.
   
   Would you like to close Chrome, Edge now? (Y/N): _
   ```

3. **Silent Mode:**
   ```powershell
   # Requires explicit flag
   Import-Bookmarks -Chrome -CloseBrowserIfRunning -TargetPath "C:\Backups"
   
   # Or use -Force
   Import-Bookmarks -Chrome -Force -TargetPath "C:\Backups"
   ```

### Closure Process

**Two-Stage Approach:**

1. **Graceful Shutdown (3 seconds):**
   - Calls `CloseMainWindow()` on each process
   - Waits up to 3 seconds for clean exit
   - Allows browser to save session data

2. **Force Kill (if needed):**
   - If process still running after 3 seconds
   - Calls `Stop-Process -Force`
   - Ensures complete closure

3. **Post-Close Delay:**
   - 2-second wait after closure
   - Allows file system to release locks
   - Ensures safe import operations

### Usage Examples

```powershell
# GUI mode: Will prompt Y/N if browsers are running
Show-BookmarkGUI
# [User imports → Browser detected → Y/N prompt → Auto-close]

# Silent mode: Auto-close without prompts
Import-Bookmarks -Chrome -Edge -CloseBrowserIfRunning -TargetPath "C:\Backups"

# Force mode: Skip all prompts
Import-Bookmarks -Firefox -Force -TargetPath "C:\Backups"

# Manual browser check
if (Test-BrowserRunning -Browser 'chrome') {
    Close-Browser -Browser 'chrome'
}
```

### Safety Features

- ✅ **Data protection**: Graceful close attempts to save sessions
- ✅ **User control**: Interactive prompts in GUI mode
- ✅ **Process verification**: Multi-process detection (helpers, renderers)
- ✅ **Retry logic**: 2-second delay before import retry
- ✅ **Logging**: All closure attempts logged
- ✅ **Fallback**: Force kill if graceful close fails

### Manual Browser Closure

If you prefer to close browsers manually:

```powershell
# Windows
Get-Process chrome,msedge,firefox | Stop-Process -Force

# Or use Task Manager (Ctrl+Shift+Esc)
```

---

## 🕑 Scheduling (Automatic Backups)

Create scheduled tasks for automatic bookmark backups.

### Create Scheduled Task

```powershell
# Default: Daily at 6:00 PM
New-BookmarkScheduledTask

# Weekly on Sundays
New-BookmarkScheduledTask -ScheduleFrequency Weekly

# Monthly on 1st day
New-BookmarkScheduledTask -ScheduleFrequency Monthly

# Custom time (Daily at 9:00 AM)
New-BookmarkScheduledTask -ScheduleFrequency Daily -Time "09:00"
```

### Task Details

- **Task Name**: `BookmarkBackupTool_AutoExport`
- **Trigger**: Based on selected frequency
- **Action**: Runs PowerShell with silent export for all browsers
- **User Context**: Runs as current user
- **Settings**: 
  - Starts even if on battery
  - Doesn't stop if going on battery
  - Starts when available (if missed)

### Command Generated

```powershell
PowerShell.exe -NoProfile -ExecutionPolicy Bypass -File "C:\Path\To\BookmarkTool.ps1" -Silent -Action Export -Chrome -Edge -Firefox
```

### View/Manage Tasks

```powershell
# View task status
Get-ScheduledTask -TaskName "BookmarkBackupTool_AutoExport"

# Run task manually
Start-ScheduledTask -TaskName "BookmarkBackupTool_AutoExport"

# Remove task
Remove-BookmarkScheduledTask
```

### Task Scheduler GUI

1. Open Task Scheduler (`taskschd.msc`)
2. Navigate to **Task Scheduler Library**
3. Find `BookmarkBackupTool_AutoExport`
4. Right-click for options (Run, Edit, Delete, etc.)

---

## ⚙️ Configuration

### Configuration File Location

Default: `%USERPROFILE%\BookmarkTool.config.json`

Custom: Use `-ConfigPath` parameter

```powershell
# Use custom config
Export-Bookmarks -Chrome -ConfigPath "C:\MyConfig\bookmarks.json"
```

### Default Configuration

```json
{
  "DefaultPath": "",
  "PreferNetworkPath": true,
  "DefaultBrowsers": ["Chrome", "Edge"],
  "AutoBackupBeforeImport": true,
  "VerifyFileIntegrity": true,
  "NetworkTimeoutSeconds": 3,
  "MaxRetryAttempts": 3,
  "RetryDelaySeconds": 1,
  "DetailedLogging": true,
  "LogRetentionDays": 30,
  "ShowProgressIndicator": true,
  "ConfirmOperations": true,
  "CreateBackupOnExport": false,
  "CompressBackups": false
}
```

### Configuration Options

| Setting                    | Type    | Default | Description                                          |
|----------------------------|---------|---------|------------------------------------------------------|
| `DefaultPath`              | String  | ""      | Default export/import path (empty = auto-detect)     |
| `PreferNetworkPath`        | Boolean | true    | Try network path (HOMESHARE) before Desktop          |
| `DefaultBrowsers`          | Array   | Chrome, Edge | Browsers to include when none specified         |
| `AutoBackupBeforeImport`   | Boolean | true    | Create backup before importing (maintains 10 copies) |
| `VerifyFileIntegrity`      | Boolean | true    | Validate JSON/SQLite structure before operations     |
| `NetworkTimeoutSeconds`    | Integer | 3       | Timeout for network path checks                      |
| `MaxRetryAttempts`         | Integer | 3       | Number of retry attempts for failed operations       |
| `RetryDelaySeconds`        | Integer | 1       | Delay between retry attempts                         |
| `DetailedLogging`          | Boolean | true    | Enable verbose logging                               |
| `LogRetentionDays`         | Integer | 30      | Days to keep log files                               |
| `ShowProgressIndicator`    | Boolean | true    | Display progress during operations                   |
| `ConfirmOperations`        | Boolean | true    | Prompt for confirmation (respects -Force)            |
| `CreateBackupOnExport`     | Boolean | false   | Create backup during export operations               |
| `CompressBackups`          | Boolean | false   | Compress backup files (separate from -CreateZip)     |

### Viewing/Modifying Configuration

```powershell
# View current config
Get-BookmarkConfiguration

# Update settings
Set-BookmarkConfiguration -AutoBackupBeforeImport $false -LogRetentionDays 60
```

### Configuration Hierarchy

1. **Command-line parameters** (highest priority)
2. **Custom config file** (via `-ConfigPath`)
3. **Default config file** (`%USERPROFILE%\BookmarkTool.config.json`)
4. **Built-in defaults** (lowest priority)

---

## 🧠 How It Works

### Export Process

1. **Initialization**
   - Load configuration
   - Check prerequisites
   - Detect browser installations

2. **Profile Detection**
   - Locate browser profile directories
   - Identify latest profile or all profiles (if `-AllProfiles`)
   - Verify bookmark files exist

3. **Path Resolution**
   - Use `-TargetPath` if specified
   - Otherwise, try network path (HOMESHARE)
   - Fallback to Desktop if network unavailable
   - Create directory if needed

4. **Export Operation**
   - Copy native bookmark files (JSON/SQLite)
   - Convert to HTML format (unless `-HtmlOnly`)
   - Apply timestamp to filenames
   - Log all operations

5. **ZIP Creation** (if `-CreateZip`)
   - Bundle all exported files
   - Create timestamped ZIP archive
   - Verify archive integrity

### Import Process

1. **Pre-Import Checks**
   - Detect running browsers
   - Prompt or auto-close (based on mode/flags)
   - Wait for processes to fully terminate (2s delay)

2. **Backup Creation** (if `AutoBackupBeforeImport`)
   - Create timestamped backup of existing bookmarks
   - Store in `<Profile>\BookmarkTool_Backups`
   - Maintain rolling 10-backup limit

3. **File Selection**
   - Use `-TargetPath` if specified
   - Detect ZIP archives and extract
   - Auto-detect browser files by naming pattern
   - Prompt for file selection (GUI mode)

4. **Integrity Validation** (if `VerifyFileIntegrity`)
   - **Chrome/Edge**: Validate JSON structure (`roots`, `version`, `bookmark_bar`)
   - **Firefox**: Verify SQLite header and database structure

5. **Import Operation**
   - Copy validated files to profile directories
   - Replace existing bookmark files
   - Log success/failure for each browser

6. **Post-Import**
   - Generate operation summary
   - Display results to user
   - Clean up temporary files (if ZIP used)

### Browser Detection Logic

**Chrome/Edge:**
- Check `%LOCALAPPDATA%\Google\Chrome\User Data` and `%LOCALAPPDATA%\Microsoft\Edge\User Data`
- Look for `Default` and `Profile *` directories
- Find directories with `Bookmarks` file
- Sort by `LastWriteTime` to identify latest

**Firefox:**
- Read `%APPDATA%\Mozilla\Firefox\profiles.ini`
- Parse `Path=` and `IsRelative=` entries
- Resolve to absolute profile path
- Verify `places.sqlite` exists
- Fallback: Search all profile directories

### Network Path Detection

1. Check `HOMESHARE` environment variable
2. Verify it's a UNC path (`\\server\share`)
3. Test accessibility with timeout (3s)
4. Create test file to verify write permissions
5. If successful, use network path
6. If failed, retry (up to 3 attempts with exponential backoff)
7. Finally, fallback to Desktop

### HTML Conversion

**Chrome/Edge:**
```
JSON → Parse → Extract bookmark_bar/other folders → 
Recursive folder traversal → Generate HTML structure → 
Write Netscape Bookmark format
```

**Firefox:**
```
SQLite → Open connection → Query moz_bookmarks + moz_places → 
Join tables → Extract URLs/titles → Generate HTML → 
Write Netscape Bookmark format
```

### Retry Logic

Operations that support retry:
- Network path access
- File copy operations
- Browser process termination
- ZIP extraction

Retry strategy:
- Initial delay: 1 second (configurable)
- Backoff multiplier: 2.0
- Max delay: 30 seconds
- Max attempts: 3 (configurable)

---

## 📂 File Locations

### Exported Files

**Standard Naming:**
```
Chrome_BookmarkData_2025-12-10_14-30-00.json
Chrome_Bookmarks_2025-12-10_14-30-00.html
Edge_BookmarkData_2025-12-10_14-30-00.json
Edge_Bookmarks_2025-12-10_14-30-00.html
Firefox_BookmarkData_2025-12-10_14-30-00.sqlite
Firefox_Bookmarks_2025-12-10_14-30-00.html
```

**Multi-Profile Naming:**
```
Chrome-Default_Profile_BookmarkData_2025-12-10_14-30-00.json
Chrome-Profile_1_Bookmarks_2025-12-10_14-30-00.html
Edge-Profile_2_BookmarkData_2025-12-10_14-30-00.json
Firefox-abc123xyz_Bookmarks_2025-12-10_14-30-00.html
```

**ZIP Archive:**
```
BookmarkBackup_2025-12-10_14-30-00.zip
```

### Tool Files

```
%USERPROFILE%\BookmarkTool.log                    # Operation logs
%USERPROFILE%\BookmarkTool.config.json            # Configuration
<ScriptDirectory>\lib\System.Data.SQLite.dll      # SQLite library
<ScriptDirectory>\lib\x64\SQLite.Interop.dll      # Native x64 DLL
<ScriptDirectory>\lib\x86\SQLite.Interop.dll      # Native x86 DLL
```

### Browser Profile Locations

**Chrome:**
```
%LOCALAPPDATA%\Google\Chrome\User Data\Default\Bookmarks
%LOCALAPPDATA%\Google\Chrome\User Data\Profile 1\Bookmarks
%LOCALAPPDATA%\Google\Chrome\User Data\Profile 2\Bookmarks
```

**Edge:**
```
%LOCALAPPDATA%\Microsoft\Edge\User Data\Default\Bookmarks
%LOCALAPPDATA%\Microsoft\Edge\User Data\Profile 1\Bookmarks
```

**Firefox:**
```
%APPDATA%\Mozilla\Firefox\Profiles\abc123xyz.default-release\places.sqlite
%APPDATA%\Mozilla\Firefox\Profiles\def456uvw.default\places.sqlite
```

### Backup Locations

```
<BrowserProfile>\BookmarkTool_Backups\Bookmarks.20251210_143000.backup
<BrowserProfile>\BookmarkTool_Backups\Bookmarks.20251210_120000.backup
<BrowserProfile>\BookmarkTool_Backups\places.sqlite.20251210_143000.backup
```

(Maintains rolling 10-backup limit)

---

## 🧭 Troubleshooting

### Common Issues

#### 1. "Browser is running" Error

**Symptom:**
```
WARNING: Chrome is currently running!
The browser must be closed to import bookmarks safely.
```

**Solutions:**

**Option A: Use GUI (Interactive)**
```powershell
Show-BookmarkGUI
# Click Import → Answer "Y" to close browser prompt
```

**Option B: Use Auto-Close Flag**
```powershell
Import-Bookmarks -Chrome -CloseBrowserIfRunning -TargetPath "C:\Backups"
```

**Option C: Manual Closure**
```powershell
# Close all browser processes
Get-Process chrome,msedge,firefox -ErrorAction SilentlyContinue | Stop-Process -Force

# Then import
Import-Bookmarks -Chrome -TargetPath "C:\Backups"
```

**Option D: Force Import**
```powershell
Import-Bookmarks -Chrome -Force -TargetPath "C:\Backups"
```

---

#### 2. "Network path not accessible" Error

**Symptom:**
```
Network path check timed out after 3 s
Using Desktop fallback: C:\Users\YourName\Desktop
```

**Solutions:**

**Option A: Verify Network Path**
```powershell
# Check HOMESHARE variable
$env:HOMESHARE

# Test access
Test-Path $env:HOMESHARE -PathType Container
```

**Option B: Use Explicit Path**
```powershell
# Specify exact path
Export-Bookmarks -Chrome -TargetPath "\\server\share\bookmarks"
```

**Option C: Adjust Timeout**
Edit `BookmarkTool.config.json`:
```json
{
  "NetworkTimeoutSeconds": 10,
  "MaxRetryAttempts": 5
}
```

**Option D: Use Local Path**
```powershell
Export-Bookmarks -Chrome -TargetPath "C:\MyBackups"
```

---

#### 3. "System.Data.SQLite not available" Warning

**Symptom:**
```
⚠ System.Data.SQLite not available - Firefox HTML conversion will be disabled
Firefox bookmarks exported as SQLite database successfully
```

**This is NOT an error** - Firefox SQLite export still works!

**To enable Firefox HTML conversion:**

**Option A: Let tool auto-download (requires internet)**
```powershell
# Tool will automatically download from NuGet on first run
Export-Bookmarks -Firefox -TargetPath "C:\Backups"
```

**Option B: Manual download**
1. Download from [NuGet: System.Data.SQLite](https://www.nuget.org/packages/Stub.System.Data.SQLite.Core.NetFramework/)
2. Extract DLLs to `<ScriptDirectory>\lib\`
3. Place `SQLite.Interop.dll` in `lib\x64\` or `lib\x86\`

**Option C: Use HTML-only export**
```powershell
# Skip SQLite file entirely
Export-Bookmarks -Firefox -HtmlOnly -TargetPath "C:\Backups"
```

---

#### 4. "Permission denied" / "Access denied" Error

**Symptom:**
```
ERROR: Cannot access target path: C:\Backups - Access denied
```

**Solutions:**

**Option A: Run as Administrator**
```powershell
# Right-click PowerShell → Run as Administrator
.\BookmarkTool.ps1
```

**Option B: Use Accessible Path**
```powershell
# Use your user profile
Export-Bookmarks -Chrome -TargetPath "$env:USERPROFILE\Documents\Bookmarks"
```

**Option C: Check Folder Permissions**
```powershell
# Verify write access
$testPath = "C:\Backups"
$testFile = Join-Path $testPath "_test.txt"
New-Item -Path $testFile -ItemType File -Force
Remove-Item $testFile -Force
```

---

#### 5. "Execution Policy" Error

**Symptom:**
```
File .\BookmarkTool.ps1 cannot be loaded. The file is not digitally signed.
```

**Solutions:**

**Option A: Bypass Policy (One-time)**
```powershell
powershell -ExecutionPolicy Bypass -File .\BookmarkTool.ps1
```

**Option B: Change Policy (Permanent)**
```powershell
# Run as Administrator
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Option C: Unblock Script**
```powershell
Unblock-File -Path .\BookmarkTool.ps1
```

---

#### 6. "Profile not found" Error

**Symptom:**
```
Chrome profile(s) not found - skipping Chrome export
```

**Solutions:**

**Check if browser is installed:**
```powershell
Test-BrowserInstalled -Browser 'Chrome'
```

**Verify profile paths:**
```powershell
# Chrome
Test-Path "$env:LOCALAPPDATA\Google\Chrome\User Data"

# Edge
Test-Path "$env:LOCALAPPDATA\Microsoft\Edge\User Data"

# Firefox
Test-Path "$env:APPDATA\Mozilla\Firefox"
```

**List available profiles:**
```powershell
Get-BrowserProfiles -Browser 'Chrome'
```

**Portable browser installations:**
- Tool only detects standard installations
- Use manual file copy for portable browsers

---

#### 7. ZIP Import/Export Issues

**Issue A: ZIP file not recognized**

```powershell
# Verify ZIP integrity
$zipPath = "C:\Backups\BookmarkBackup.zip"
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::OpenRead($zipPath)
```

**Issue B: Extraction fails**

```powershell
# Check available space
Get-PSDrive C | Select-Object Used,Free

# Try manual extraction
Expand-Archive -Path "C:\Backups\BookmarkBackup.zip" -DestinationPath "C:\Temp\BookmarkExtract"
```

**Issue C: Auto-detection fails**

```powershell
# Import specific browser from ZIP
Import-Bookmarks -Chrome -TargetPath "C:\Backups\BookmarkBackup.zip"
```

---

#### 8. HTML Conversion Produces Empty Files

**Symptom:**
HTML file is created but contains no bookmarks

**Solutions:**

**Check source file integrity:**
```powershell
Test-BookmarkFileIntegrity -FilePath "C:\...\Bookmarks" -BrowserType 'chrome'
```

**Verify bookmark data exists:**
```powershell
# For Chrome/Edge
$json = Get-Content "C:\...\Bookmarks" -Raw | ConvertFrom-Json
$json.roots.bookmark_bar.children.Count

# For Firefox - requires SQLite
```

**Enable verbose logging:**
```powershell
Export-Bookmarks -Chrome -Verbose -TargetPath "C:\Backups"
```

---

#### 9. Log File Issues

**Log file grows too large:**

```powershell
# Adjust retention in config
Set-BookmarkConfiguration -LogRetentionDays 7
```

**Clear log manually:**

```powershell
Remove-Item "$env:USERPROFILE\BookmarkTool.log" -Force
```

**View recent log entries:**

```powershell
Get-Content "$env:USERPROFILE\BookmarkTool.log" -Tail 50
```

---

### Advanced Troubleshooting

#### Enable Detailed Logging

```powershell
# Set in config
Set-BookmarkConfiguration -DetailedLogging $true

# Or use -Verbose flag
Export-Bookmarks -Chrome -Verbose
```

#### Check Prerequisites

```powershell
Test-BookmarkPrerequisites -TargetPath "C:\Backups"
```

#### Test Individual Components

```powershell
# Test browser detection
Test-BrowserRunning -Browser 'chrome'

# Test profile detection
Get-BrowserProfiles -Browser 'Chrome'

# Test network path
Get-HomeSharePath

# Test file integrity
Test-BookmarkFileIntegrity -FilePath "C:\...\Bookmarks" -BrowserType 'chrome'
```

#### Retry Failed Operations

```powershell
# Operations automatically retry 3 times
# Adjust retry settings in config:
Set-BookmarkConfiguration -MaxRetryAttempts 5 -RetryDelaySeconds 2
```

---

## ❓ FAQ

### General Questions

**Q: Does this tool merge bookmarks or replace them?**  
A: **Replaces.** The tool replaces the entire bookmark file. If you need to merge, export first, then manually combine in the browser.

**Q: Can I use this with portable browser versions?**  
A: Portable browsers store data in non-standard locations. You'll need to manually specify paths using `-TargetPath`.

**Q: Does this work with browser sync enabled?**  
A: Yes, but imported bookmarks will sync to your account. Be cautious when importing to synced profiles.

**Q: Can I schedule backups for multiple users?**  
A: Yes, but each user needs their own scheduled task. Loop through `C:\Users` for automation:
```powershell
Get-ChildItem C:\Users | ForEach-Object {
    $userProfile = $_.FullName
    # Create task for each user
}
```

---

### Export Questions

**Q: What's the difference between native and HTML export?**  
A:
- **Native**: Browser-specific format (JSON/SQLite), fastest to import back
- **HTML**: Universal format, works with any browser, human-readable

**Q: Should I use `-HtmlOnly` or default dual-format export?**  
A:
- **Dual-format (default)**: Best for complete backups
- **HTML-only**: Best for cross-browser sharing, archival, smaller file sizes

**Q: How much disk space do exports use?**  
A: Typical sizes:
- Chrome/Edge JSON: 50KB - 5MB
- Firefox SQLite: 5MB - 50MB
- HTML files: 20KB - 2MB
- ZIP archive: 50% smaller than combined files

**Q: Can I export bookmarks while the browser is running?**  
A: **Yes for exports, no for imports.** Exports can run while browser is open. Imports require browser to be closed.

**Q: What happens if I export with `-AllProfiles`?**  
A: Creates separate files for each profile:
```
Chrome-Default_Profile_BookmarkData_timestamp.json
Chrome-Profile_1_BookmarkData_timestamp.json
Chrome-Profile_2_BookmarkData_timestamp.json
```

---

### Import Questions

**Q: What happens to my existing bookmarks during import?**  
A: They're backed up first (if `AutoBackupBeforeImport` is true), then replaced. Check `<Profile>\BookmarkTool_Backups` for backups.

**Q: Can I restore a specific backup?**  
A: Yes, manually:
```powershell
$backup = "C:\...\Profile\BookmarkTool_Backups\Bookmarks.20251210_143000.backup"
$target = "C:\...\Profile\Bookmarks"
Copy-Item $backup $target -Force
```

**Q: How many automatic backups are kept?**  
A: 10 rolling backups per browser profile. Oldest are automatically deleted.

**Q: What if the browser won't close?**  
A: The tool tries:
1. Graceful close (3s wait)
2. Force kill
3. If both fail, manually close with Task Manager

**Q: Can I import bookmarks to a different computer?**  
A: Yes! Export on computer A, transfer files, then import on computer B:
```powershell
# Computer A
Export-Bookmarks -Chrome -CreateZip -TargetPath "D:\USB"

# Computer B
Import-Bookmarks -Chrome -TargetPath "D:\USB\BookmarkBackup_*.zip"
```

---

### ZIP Archive Questions

**Q: When should I use `-CreateZip`?**  
A: Use ZIP archives for:
- Email transfers
- USB backup
- Cloud storage
- Long-term archival
- Multiple browser backups in one file

**Q: Can I import from a ZIP without extracting it first?**  
A: Yes! The tool automatically extracts and imports:
```powershell
Import-Bookmarks -Chrome -Edge -TargetPath "C:\Backups\BookmarkBackup.zip"
```

**Q: What's inside the ZIP file?**  
A: All exported bookmark files (native + HTML) for selected browsers.

---

### HTML Conversion Questions

**Q: Why does Firefox HTML conversion require SQLite?**  
A: Firefox stores bookmarks in a SQLite database. System.Data.SQLite is needed to query the database and extract bookmarks.

**Q: What if SQLite auto-download fails?**  
A: Firefox exports still work (SQLite file is exported). Only HTML conversion is disabled. Manual SQLite installation is possible.

**Q: Are HTML bookmarks compatible with all browsers?**  
A: Yes! Netscape Bookmark format is universally supported:
- Chrome: Import via `chrome://bookmarks` → ⋮ → Import
- Edge: Settings → Profiles → Import browser data
- Firefox: Bookmarks → Show All → Import and Backup → Import

**Q: Can I edit HTML bookmark files?**  
A: Yes, they're plain text. Be careful with the structure:
```html
<DT><A HREF="https://example.com">Bookmark Title</A>
```

---

### Scheduling Questions

**Q: How do I change the scheduled backup time?**  
A:
```powershell
# Remove existing task
Remove-BookmarkScheduledTask

# Create new task with different time
New-BookmarkScheduledTask -ScheduleFrequency Daily -Time "09:00"
```

**Q: Do scheduled backups use ZIP archives?**  
A: No, by default. To enable, edit the scheduled task in Task Scheduler and add `-CreateZip` to the arguments.

**Q: Where do scheduled backups save files?**  
A: To the auto-detected path (network share or Desktop). Specify with config:
```json
{
  "DefaultPath": "C:\\MyBackups"
}
```

**Q: Will scheduled tasks work if I'm not logged in?**  
A: They run when you're logged in. For always-on backups, set task to "Run whether user is logged on or not" in Task Scheduler (requires admin).

---

### Troubleshooting Questions

**Q: Why can't the tool find my browser?**  
A: Check:
1. Browser installed in standard location?
2. Portable installation? (not supported without `-TargetPath`)
3. Custom profile location? (use `-AllProfiles`)

**Q: Network path detection always fails?**  
A: Check:
1. `$env:HOMESHARE` is set and valid
2. Network is accessible
3. Write permissions exist
4. Increase timeout in config (`NetworkTimeoutSeconds`)

**Q: Import doesn't seem to work?**  
A: Verify:
1. Browser was closed during import
2. File integrity passed (`Test-BookmarkFileIntegrity`)
3. Check log file: `$env:USERPROFILE\BookmarkTool.log`
4. Restart browser to see imported bookmarks

---

### Configuration Questions

**Q: Where is the config file stored?**  
A: `%USERPROFILE%\BookmarkTool.config.json` (usually `C:\Users\YourName\BookmarkTool.config.json`)

**Q: Can I use different configs for different scenarios?**  
A: Yes:
```powershell
Export-Bookmarks -Chrome -ConfigPath "C:\Configs\work_bookmarks.json"
Export-Bookmarks -Firefox -ConfigPath "C:\Configs\personal_bookmarks.json"
```

**Q: What happens if the config file is missing?**  
A: Built-in defaults are used automatically.

---

### Performance Questions

**Q: How long does export take?**  
A: Typically 1-5 seconds per browser. Factors:
- Bookmark count (more = longer)
- Network path (slower than local)
- HTML conversion (adds 1-2 seconds)
- ZIP creation (adds 1-2 seconds)

**Q: Can I speed up exports?**  
A: Yes:
- Use `-HtmlOnly` (skips native file copy)
- Use local path instead of network
- Skip `-CreateZip`
- Disable integrity checks in config

**Q: Do imports slow down my browser?**  
A: No. Browser reads bookmarks on startup. No performance impact.

---

## 🔐 Security Notes

### Data Protection

**Bookmark Content:**
- ✅ Bookmarks may contain sensitive URLs (internal sites, personal accounts)
- ✅ Treat bookmark backups as confidential data
- ✅ Encrypt sensitive backups (use BitLocker, 7-Zip encryption, etc.)
- ✅ Secure network shares with appropriate ACLs

**File Permissions:**
- ✅ Exported files inherit directory permissions
- ✅ Verify backup location has restricted access
- ✅ Remove backups from shared drives after transfer

### Network Path Security

**HOMESHARE Variable:**
- ✅ Ensure `HOMESHARE` points to trusted, secured network location
- ✅ Use authentication-required shares (avoid anonymous access)
- ✅ Monitor network share access logs

### PowerShell Execution

**Script Execution:**
- ✅ Script is not code-signed (requires execution policy adjustment)
- ✅ Review script contents before first run
- ✅ Use `Get-FileHash` to verify script integrity if downloaded
- ✅ Consider code-signing for enterprise deployment

**Execution Policy:**
```powershell
# View current policy
Get-ExecutionPolicy

# Set policy (user-specific)
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

# Bypass for single run (safest)
powershell -ExecutionPolicy Bypass -File .\BookmarkTool.ps1
```

### Auto-Downloaded Components

**System.Data.SQLite:**
- ✅ Downloaded from official NuGet.org
- ✅ Uses HTTPS (TLS 1.2+)
- ✅ Package: Stub.System.Data.SQLite.Core.NetFramework v1.0.118
- ✅ Installed locally to script directory (`lib` folder)
- ⚠️ Review NuGet package before auto-download in enterprise environments

### Scheduled Tasks

**Task Security:**
- ✅ Tasks run in user context (not SYSTEM)
- ✅ No elevation required for scheduled backups
- ✅ Task definition is visible in Task Scheduler
- ✅ Review task arguments before approving

### Browser Auto-Close

**Process Termination:**
- ✅ Attempts graceful close first (safe)
- ⚠️ May force-kill if needed (potential data loss if unsaved forms)
- ✅ Only closes browser processes, not other apps
- ✅ User confirmation required in GUI mode

### Log Files

**Logging Information:**
- ⚠️ Logs may contain file paths, usernames, profile names
- ✅ No passwords or sensitive bookmark URLs logged
- ✅ Logs stored in user profile: `%USERPROFILE%\BookmarkTool.log`
- ✅ Automatic retention cleanup after 30 days (configurable)

**Secure Logging:**
```powershell
# Adjust retention
Set-BookmarkConfiguration -LogRetentionDays 7

# Manually clear logs
Remove-Item "$env:USERPROFILE\BookmarkTool.log" -Force
```

### Enterprise Deployment

**Group Policy Recommendations:**
- ✅ Deploy via SCCM/Intune with approved digital signature
- ✅ Use centralized config file (map `-ConfigPath` to network location)
- ✅ Restrict backup locations to approved network shares
- ✅ Audit scheduled task creation
- ✅ Monitor bookmark exports to detect data exfiltration

**AppLocker / Script Policies:**
- ✅ Whitelist script by hash or path
- ✅ Restrict NuGet auto-download in air-gapped environments
- ✅ Pre-install System.Data.SQLite in locked-down environments

### Best Practices

1. **Verify Script Source**
   ```powershell
   # Get file hash
   Get-FileHash .\BookmarkTool.ps1 -Algorithm SHA256
   ```

2. **Use Secure Backup Locations**
   ```powershell
   # Encrypted local drive
   Export-Bookmarks -Chrome -TargetPath "E:\SecureBackups"
   
   # Authenticated network share
   Export-Bookmarks -Chrome -TargetPath "\\secure-server\backups\$env:USERNAME"
   ```

3. **Regular Backup Rotation**
   ```powershell
   # Create timestamped backups
   Export-Bookmarks -Chrome -CreateZip -TargetPath "C:\Backups"
   
   # Keep only last 5 archives
   Get-ChildItem "C:\Backups\BookmarkBackup_*.zip" | 
     Sort-Object LastWriteTime -Descending | 
     Select-Object -Skip 5 | 
     Remove-Item -Force
   ```

4. **Audit Trail**
   ```powershell
   # Review recent operations
   Get-Content "$env:USERPROFILE\BookmarkTool.log" -Tail 100
   
   # Filter for specific actions
   Select-String -Path "$env:USERPROFILE\BookmarkTool.log" -Pattern "Import|Export"
   ```

5. **Least Privilege**
   - Run as standard user (not admin) when possible
   - Use admin only for scheduled task creation
   - Restrict file permissions on backup directories

---

## ⚠️ Known Limitations

### Functional Limitations

**1. No Bookmark Merging**
- **Limitation**: Import replaces entire bookmark file
- **Impact**: Cannot merge new bookmarks with existing ones
- **Workaround**: 
  - Export before import (creates backup)
  - Manually merge in browser UI
  - Use browser's built-in import to merge

**2. No Cloud Sync Detection**
- **Limitation**: Tool doesn't detect if browser sync is enabled
- **Impact**: Imported bookmarks will sync to cloud account
- **Workaround**: 
  - Disable sync before import
  - Be aware imports affect all synced devices

**3. Portable Browser Support**
- **Limitation**: Only detects standard installation paths
- **Impact**: Portable browsers not auto-detected
- **Workaround**:
  ```powershell
  # Manually specify portable browser path
  Export-Bookmarks -Chrome -TargetPath "D:\PortableChrome\Profile\Bookmarks"
  ```

**4. Firefox Multi-Container Support**
- **Limitation**: HTML export doesn't preserve container assignments
- **Impact**: Firefox containers (work/personal) lost in HTML format
- **Workaround**: Use native SQLite format for full Firefox features

**5. Bookmark Folder Permissions**
- **Limitation**: Doesn't preserve Chrome's managed bookmarks or enterprise policies
- **Impact**: Policy-enforced bookmarks may reappear after import
- **Workaround**: Re-apply Group Policy after import

---

### Technical Limitations

**6. Windows-Only**
- **Limitation**: Requires Windows 10/11, .NET Framework
- **Impact**: No macOS/Linux support in PowerShell version
- **Workaround**: Use dedicated [macOS app](https://github.com/hov172/MacOS-Bookmarks-Backup-Tool)

**7. PowerShell 5.1+ Required**
- **Limitation**: Older PowerShell versions not supported
- **Impact**: Windows 7/8 may have compatibility issues
- **Workaround**: Update PowerShell or use WMF 5.1

**8. SQLite Architecture Dependency**
- **Limitation**: System.Data.SQLite requires matching architecture (x86/x64)
- **Impact**: Auto-download may fail on ARM64 systems
- **Workaround**: Firefox exports still work (SQLite file only, no HTML)

**9. Network Path Timeout**
- **Limitation**: 3-second timeout for network path checks
- **Impact**: Slow networks may fallback to Desktop unnecessarily
- **Workaround**:
  ```json
  {
    "NetworkTimeoutSeconds": 10,
    "MaxRetryAttempts": 5
  }
  ```

**10. UNC Path Limitations**
- **Limitation**: SQLite cannot open databases on UNC paths
- **Impact**: Firefox HTML conversion may fail on network shares
- **Workaround**: Tool automatically copies to temp location

---

### GUI Limitations

**11. STA Thread Requirement**
- **Limitation**: GUI requires Single-Threaded Apartment (STA) mode
- **Impact**: Auto-relaunches if not in STA (slight delay)
- **Workaround**: None needed (automatic)

**12. Windows Forms Dependency**
- **Limitation**: Requires System.Windows.Forms assembly
- **Impact**: Won't work in PowerShell Core (pwsh) without Windows
- **Workaround**: Use Windows PowerShell 5.1

**13. No Drag-and-Drop**
- **Limitation**: Cannot drag files into GUI
- **Impact**: Must use Browse button
- **Workaround**: Use CLI mode with `-TargetPath`

---

### Browser-Specific Limitations

**14. Chrome/Edge Profile Detection**
- **Limitation**: Relies on standard `User Data` folder structure
- **Impact**: Custom profile locations not detected
- **Workaround**: Use `-TargetPath` to specify custom profiles

**15. Firefox Profile Selection**
- **Limitation**: Detects primary profile from `profiles.ini`
- **Impact**: May not detect all profiles if `profiles.ini` is corrupted
- **Workaround**: Use `-AllProfiles` or manually specify profile path

**16. Browser Version Compatibility**
- **Limitation**: Tested with Chrome 100+, Edge 100+, Firefox 100+
- **Impact**: Very old browser versions may have different file formats
- **Workaround**: Update browsers to latest versions

**17. Brave/Vivaldi/Opera Support**
- **Limitation**: Chromium-based browsers not explicitly supported
- **Impact**: May work but not guaranteed
- **Workaround**:
  ```powershell
  # Try using Chrome export on Brave profile
  $bravePath = "$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data\Default"
  Export-Bookmarks -Chrome -TargetPath $bravePath
  ```

---

### Scheduling Limitations

**18. Monthly Schedule Accuracy**
- **Limitation**: Monthly schedule uses 30-day interval (not calendar month)
- **Impact**: May drift slightly over time
- **Workaround**: Manually adjust task in Task Scheduler

**19. Missed Scheduled Runs**
- **Limitation**: If computer is off during scheduled time
- **Impact**: Backup is skipped (unless "Start when available" is set)
- **Workaround**: Task is configured with "Start when available" by default

**20. Multi-User Scheduling**
- **Limitation**: One scheduled task per user session
- **Impact**: Must create separate tasks for each user account
- **Workaround**:
  ```powershell
  # Create task for multiple users
  foreach ($user in Get-LocalUser) {
      # Create task with $user credentials
  }
  ```

---

### File Format Limitations

**21. HTML Format Fidelity**
- **Limitation**: HTML export doesn't preserve all metadata:
  - Chrome: Missing favicons, date modified
  - Firefox: Missing tags, keywords, container assignments
- **Impact**: Some bookmark metadata lost in HTML format
- **Workaround**: Use native format (JSON/SQLite) for complete backups

**22. ZIP Compression Ratio**
- **Limitation**: Text files (JSON/HTML) compress well (~50%), SQLite databases less so (~10-20%)
- **Impact**: Firefox ZIP archives are large
- **Workaround**: Use external compression tools for better ratios

**23. File Size Limits**
- **Limitation**: No built-in file size limits
- **Impact**: Extremely large bookmark collections (1M+ bookmarks) may be slow
- **Workaround**: Process individual profiles separately with `-AllProfiles`

---

### Security Limitations

**24. No Encryption**
- **Limitation**: Exported files are not encrypted
- **Impact**: Bookmarks stored in plain text
- **Workaround**: Use BitLocker, 7-Zip encryption, or encrypted network shares

**25. No Digital Signature**
- **Limitation**: Script is not code-signed
- **Impact**: Requires execution policy adjustment
- **Workaround**: Sign script for enterprise deployment

**26. No File Integrity Verification**
- **Limitation**: No cryptographic hash verification of backups
- **Impact**: Cannot detect corrupted or tampered files
- **Workaround**: Use `Get-FileHash` manually

---

### Enterprise Limitations

**27. No Group Policy Integration**
- **Limitation**: Cannot be configured via Group Policy
- **Impact**: Must deploy config files manually
- **Workaround**: Use network-shared config file with `-ConfigPath`

**28. No SCCM/Intune Packages**
- **Limitation**: No pre-built MSI or MSIX packages
- **Impact**: Manual deployment required
- **Workaround**: Wrap in custom MSI or use PowerShell App Deployment Toolkit

**29. No Centralized Logging**
- **Limitation**: Logs stored locally per user
- **Impact**: No central audit trail
- **Workaround**: Configure network log path in `DefaultPath`

**30. No User-Level Encryption**
- **Limitation**: Cannot enforce encryption on backups
- **Impact**: Users may store backups insecurely
- **Workaround**: Use Group Policy to restrict backup locations to encrypted drives

---

### Future Enhancements Under Consideration

- ✨ Bookmark merging capability
- ✨ Differential/incremental backups
- ✨ Native macOS/Linux support in PowerShell Core
- ✨ Cloud storage integration (OneDrive, Dropbox, Google Drive)
- ✨ Automatic backup before browser updates
- ✨ Bookmark validation (broken link detection)
- ✨ Encrypted backup option
- ✨ GUI modernization (WPF or web-based)
- ✨ Brave/Vivaldi/Opera explicit support
- ✨ Firefox container preservation in HTML export

---

## 🧹 Uninstall / Cleanup

### Remove Module

```powershell
# Uninstall PowerShell module
Uninstall-Module -Name BookmarkBackupTool -Force

# Verify removal
Get-Module -ListAvailable -Name BookmarkBackupTool
```

### Remove Scheduled Tasks

```powershell
# Remove scheduled task
Remove-BookmarkScheduledTask

# Or manually via Task Scheduler
# Open taskschd.msc → Delete "BookmarkBackupTool_AutoExport"
```

### Remove Configuration and Logs

```powershell
# Remove config file
Remove-Item "$env:USERPROFILE\BookmarkTool.config.json" -Force -ErrorAction SilentlyContinue

# Remove log file
Remove-Item "$env:USERPROFILE\BookmarkTool.log" -Force -ErrorAction SilentlyContinue

# Remove SQLite library (if auto-downloaded)
$scriptPath = Split-Path $PSCommandPath -Parent
Remove-Item "$scriptPath\lib" -Recurse -Force -ErrorAction SilentlyContinue
```

### Remove Bookmark Backups

```powershell
# Chrome
Remove-Item "$env:LOCALAPPDATA\Google\Chrome\User Data\*\BookmarkTool_Backups" -Recurse -Force -ErrorAction SilentlyContinue

# Edge
Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\BookmarkTool_Backups" -Recurse -Force -ErrorAction SilentlyContinue

# Firefox
Remove-Item "$env:APPDATA\Mozilla\Firefox\Profiles\*\BookmarkTool_Backups" -Recurse -Force -ErrorAction SilentlyContinue
```

### Complete Cleanup Script

```powershell
# Complete removal of all Bookmark Tool components
Write-Host "Removing Bookmark Backup Tool..." -ForegroundColor Yellow

# 1. Remove module
if (Get-Module -ListAvailable -Name BookmarkBackupTool) {
    Uninstall-Module -Name BookmarkBackupTool -Force
    Write-Host "✓ Module removed" -ForegroundColor Green
}

# 2. Remove scheduled task
try {
    Unregister-ScheduledTask -TaskName "BookmarkBackupTool_AutoExport" -Confirm:$false -ErrorAction Stop
    Write-Host "✓ Scheduled task removed" -ForegroundColor Green
} catch {
    Write-Host "  Scheduled task not found (already removed)" -ForegroundColor Gray
}

# 3. Remove config and logs
Remove-Item "$env:USERPROFILE\BookmarkTool.config.json" -Force -ErrorAction SilentlyContinue
Remove-Item "$env:USERPROFILE\BookmarkTool.log" -Force -ErrorAction SilentlyContinue
Write-Host "✓ Config and logs removed" -ForegroundColor Green

# 4. Remove SQLite library
$scriptPath = $PSScriptRoot
if (Test-Path "$scriptPath\lib") {
    Remove-Item "$scriptPath\lib" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "✓ SQLite library removed" -ForegroundColor Green
}

# 5. Remove bookmark backups (optional - prompts user)
$removeBackups = Read-Host "Remove automatic bookmark backups from browser profiles? (Y/N)"
if ($removeBackups -match '^[Yy]') {
    Remove-Item "$env:LOCALAPPDATA\Google\Chrome\User Data\*\BookmarkTool_Backups" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:LOCALAPPDATA\Microsoft\Edge\User Data\*\BookmarkTool_Backups" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item "$env:APPDATA\Mozilla\Firefox\Profiles\*\BookmarkTool_Backups" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "✓ Bookmark backups removed" -ForegroundColor Green
} else {
    Write-Host "  Bookmark backups preserved" -ForegroundColor Gray
}

Write-Host "`nBookmark Backup Tool completely removed." -ForegroundColor Green
Write-Host "Your browser bookmarks were not affected." -ForegroundColor Cyan
```

### Preserve User Data

**To keep your exported bookmarks safe during uninstall:**

```powershell
# User's exported bookmarks are NOT removed by uninstall
# They remain in whatever location you exported them to:
# - Desktop
# - Network share (HOMESHARE)
# - Custom paths (-TargetPath)

# Example locations:
# C:\Users\YourName\Desktop\Chrome_BookmarkData_*.json
# \\server\share\BookmarkBackup_*.zip
```

### Reinstall

```powershell
# Reinstall fresh copy
Install-Module -Name BookmarkBackupTool -Scope CurrentUser -Force

# Or update if already installed
Update-Module -Name BookmarkBackupTool
```

---

## 📜 License

[MIT License](https://opensource.org/license/mit/)

Copyright (c) 2025 Jesus M. Ayala

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

---

## 🙌 Credits

**Author:** Jesus M. Ayala

**Version:** 5.2 Enhanced Edition

**Last Updated:** December 10, 2025

**GitHub Repositories:**
- 💻 [MacOS Bookmarks Backup Tool](https://github.com/hov172/MacOS-Bookmarks-Backup-Tool)
- 🖥️ [Windows Bookmarks Backup Tool](https://github.com/hov172/Win-Bookmarks-Backup-Tool/tree/main)

**PowerShell Gallery:**
- 📦 [BookmarkBackupTool Module](https://www.powershellgallery.com/packages/BookmarkBackupTool)

**Special Thanks:**
- Microsoft PowerShell Team
- System.Data.SQLite Project
- Netscape Bookmark File Format Specification
- Community testers and contributors

---

## 📞 Support & Contributions

**Report Issues:**
- GitHub Issues: [Win-Bookmarks-Backup-Tool/issues](https://github.com/hov172/Win-Bookmarks-Backup-Tool/issues)

**Feature Requests:**
- Submit via GitHub Issues with `[Feature Request]` tag

**Contributions:**
- Pull requests welcome!
- Follow existing code style
- Include tests for new features
- Update documentation

**Questions:**
- Check [FAQ](#-faq) first
- Review [Troubleshooting](#-troubleshooting)
- Open GitHub Discussion for general questions

---

**⭐ If you find this tool useful, please star the repository!**

**📢 Share with others who might benefit from automated bookmark backups!**

---

**End of README**
