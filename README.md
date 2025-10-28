# 📑 Bookmark Backup Tool v5.0 — PowerShell Enhanced Edition

PowerShell tool that lets you **export and import bookmarks** for **Google Chrome**, **Microsoft Edge**, and **Mozilla Firefox**.  
It supports GUI for interactive use and CLI for automation, offers scheduled backups, multi-profile handling, file integrity checks, and comprehensive logging.  
It's designed to be **robust in enterprise environments** (network paths, permissions, and process safety) while staying simple for individual users.

**I built both Windows and macOS applications based on this script. These apps allow users to manage their own backups, and both versions include command-line options for administrators.**

- 💻 **MacOS**: [https://github.com/hov172/MacOS-Bookmarks-Backup-Tool](https://github.com/hov172/MacOS-Bookmarks-Backup-Tool)  
- 🖥️ **Windows**: [https://github.com/hov172/Win-Bookmarks-Backup-Tool/tree/main](https://github.com/hov172/Win-Bookmarks-Backup-Tool/tree/main)

---

<img width="579" height="410" alt="bookmark" src="https://github.com/user-attachments/assets/f9c5ad00-9372-4d14-9c68-5846da5c95fe" />

---

## 📝 Changelog
**v5.0 — October 25, 2025**

- Made into a Powershell Module
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

**Bookmark Backup Tool v5.0** is a PowerShell module that lets you export and import bookmarks for **Google Chrome**, **Microsoft Edge**, and **Mozilla Firefox**.  
It supports **GUI** for interactive use and **CLI** for automation, offers **scheduled backups**, **multi-profile handling**, **file integrity checks**, and **comprehensive logging**.  
It's built to work smoothly in both **enterprise** and **home** environments.

---

## ✨ Key Features

- Multi-browser support (Chrome, Edge, Firefox)
- GUI (Windows Forms) and CLI (silent mode)
- Automatic path detection with network share preference and Desktop fallback
- Multi-profile support
- Automatic backup before imports
- File integrity verification for JSON and SQLite files
- Task Scheduler integration (Daily/Weekly/Monthly)
- Configuration file support
- Verbose logging with retention policy
- Safety mechanisms: browser process detection, retry logic, WhatIf/Force switches
- PowerShell Module

---

## 🧰 Requirements

- Windows 10 or 11  
- PowerShell 5.1+  
- .NET Assemblies: `System.Windows.Forms`, `System.Drawing`  
- File system access to profiles/target paths  
- Admin rights for scheduled tasks (sometimes required)

---

## 🌐 Supported Browsers & Files

| Browser  | Exported/Imported File         | Native Location                                                                 |
|----------|-------------------------------|----------------------------------------------------------------------------------|
| Chrome   | `Chrome-Bookmarks.json`       | `%LOCALAPPDATA%\Google\Chrome\User Data\<Profile>\Bookmarks`                    |
| Edge     | `Edge-Bookmarks.json`         | `%LOCALAPPDATA%\Microsoft\Edge\User Data\<Profile>\Bookmarks`                   |
| Firefox  | `Firefox-places.sqlite`       | `%APPDATA%\Mozilla\Firefox\Profiles\<Profile>\places.sqlite`                    |

---

## 📦 Installation (PowerShell Gallery)

👉 [PowerShell Gallery Package](https://www.powershellgallery.com/packages/BookmarkBackupTool/5.0.0)

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
| `Export-Bookmarks`              | Exports bookmarks from supported browsers to a backup file.           |
| `Import-Bookmarks`              | Imports bookmarks from a backup file.                                 |
| `Get-HomeSharePath`            | Returns the user's home or shared folder path.                        |
| `Test-BrowserInstalled`        | Checks if a supported browser is installed.                           |
| `Test-BrowserRunning`          | Checks if a browser is currently running.                             |
| `Get-BrowserProfiles`          | Lists available profiles.                                            |
| `Backup-ExistingBookmarks`     | Creates a snapshot of the current bookmarks.                           |
| `Show-BookmarkGUI`             | Launches the graphical interface.                                     |
| `New-BookmarkScheduledTask`    | Creates a scheduled task for automated backups.                       |
| `Remove-BookmarkScheduledTask` | Removes a scheduled backup task.                                      |
| `Get-BookmarkConfiguration`    | Displays current configuration.                                       |
| `Set-BookmarkConfiguration`    | Updates configuration.                                               |
| `Test-BookmarkPrerequisites`   | Checks environment readiness.                                        |

To list all commands:

```powershell
Get-Command -Module BookmarkBackupTool
```

---

## ⚡ Quick Start

```powershell
# Launch GUI
Show-BookmarkGUI
```

Or export bookmarks directly:

```powershell
Export-Bookmarks -Browser Chrome -Path "$env:USERPROFILE\Documents\chrome-bookmarks.json"
```

---

## 🖱️ Usage

### GUI Mode (Beginner-friendly)

```powershell
Show-BookmarkGUI
```

### CLI / Silent Mode

Requires `-Silent` and `-Action Export|Import` with browser flags.

---

### 🧾 Parameters

| Parameter               | Description                                  |
|-------------------------|-----------------------------------------------|
| `-Silent`               | Run without GUI                              |
| `-Action`               | `Export` or `Import`                          |
| `-Chrome` / `-Edge` / `-Firefox` | Select browsers                      |
| `-TargetPath`           | Override detected path                        |
| `-WhatIf`               | Preview only                                 |
| `-Force`                | Force without confirmation                    |
| `-AllProfiles`          | Process all profiles                          |
| `-CreateScheduledTask`  | Create auto-backup task                        |
| `-ScheduleFrequency`    | Daily/Weekly/Monthly                          |
| `-ConfigPath`           | Use custom config file                         |

---

### 🧰 Command-Line Examples

#### Export
```powershell
Export-Bookmarks -Browser Chrome -Path "D:\Backups\Chrome.json"
```

#### Import
```powershell
Import-Bookmarks -Browser Edge -Path "D:\Backups\Edge.json"
```

#### Schedule
```powershell
New-BookmarkScheduledTask -Browser Firefox -Path "D:\Backups\firefox.json" -ScheduleFrequency Daily
```

#### Check environment
```powershell
Test-BookmarkPrerequisites
```

---

## 🕑 Scheduling (Automatic Backups)

- Daily, Weekly, Monthly options (default 6:00 PM)  
- Creates a Task Scheduler job named `BookmarkBackupTool_AutoExport`.

---

## ⚙️ Configuration

Default config path:
```
%USERPROFILE%\BookmarkTool.config.json
```

Example:

```json
{
  "PreferNetworkPath": true,
  "DefaultBrowsers": ["Chrome","Edge"],
  "VerifyFileIntegrity": true,
  "AutoBackupBeforeImport": true
}
```

---

## 🧠 How It Works

- Detects browser profiles (Chrome/Edge/Firefox)  
- Uses network paths if available, otherwise Desktop fallback  
- Integrity checks and safe backups before overwrite  
- Detailed logging with retention policy

---

## 📂 File Locations

- Exports: `Chrome-Bookmarks.json`, `Edge-Bookmarks.json`, `Firefox-places.sqlite`  
- Logs: `%USERPROFILE%\BookmarkTool.log`  
- Config: `%USERPROFILE%\BookmarkTool.config.json`  
- Auto backups: `<Profile>\BookmarkTool_Backups`

---

## 🧭 Troubleshooting

- **Browser is running** → Close browsers or run:
  ```powershell
  Get-Process chrome,msedge,firefox | Stop-Process
  ```
- **Network path error** → Fallback to Desktop or specify `-TargetPath`  
- **Permission denied** → Run PowerShell as Administrator  
- **Execution policy**:
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\BookMarkToolv5.ps1
  ```

---

## ❓ FAQ

- Does this merge bookmarks? → ❌ No, it replaces files.  
- Can I compress backups? → ✅ Yes, use `Compress-Archive` after export.  
- Can I automate multiple users? → ✅ Yes, loop through `C:\Users`.

---

## 🔐 Security Notes

- Exported bookmarks may contain sensitive data.  
- Use secure storage and protect network shares.

---

## ⚠️ Known Limitations

- No merge capability.  
- Monthly schedule = 4-week interval.  
- Enterprise profile redirection may affect detection.

---

## 🧹 Uninstall / Cleanup

```powershell
Unregister-ScheduledTask -TaskName "BookmarkBackupTool_AutoExport" -Confirm:$false
Remove-Item "$env:USERPROFILE\BookmarkTool.config.json" -ErrorAction SilentlyContinue
Remove-Item "$env:USERPROFILE\BookmarkTool.log" -ErrorAction SilentlyContinue
```

To remove the module:
```powershell
Uninstall-Module -Name BookmarkBackupTool
```

---

## 📜 License

[MIT License](https://opensource.org/license/mit/)

---

## 🙌 Credits

**Author:** Jesus M. Ayala

---
