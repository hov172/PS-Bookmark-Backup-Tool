# Bookmark Backup Tool v5.0 -- Powershell Enhanced Edition

PowerShell tool that lets you export and import bookmarks for Google Chrome, Microsoft Edge, and Mozilla Firefox. It supports GUI for interactive use and CLI for automation, offers scheduled backups, multi-profile handling, file integrity checks, and comprehensive logging. It’s designed to be robust in enterprise environments (network paths, permissions, and process safety) while staying simple for individual users.

** I built both Windows and MacOS applications based on this script. These apps will allow users to manage their own backups. Both versions will also include command-line options for administrators. 

- MacOS: https://github.com/hov172/MacOS-Bookmarks-Backup-Tool
- Windows: https://github.com/hov172/Win-Bookmarks-Backup-Tool/tree/main
------------------------------------------------------------------------
<img width="579" height="410" alt="bookmark" src="https://github.com/user-attachments/assets/f9c5ad00-9372-4d14-9c68-5846da5c95fe" />

------------------------------------------------------------------------
## Table of Contents

-   [Overview](#overview)
-   [Key Features](#key-features)
-   [Requirements](#requirements)
-   [Supported Browsers & Files](#supported-browsers--files)
-   [Quick Start](#quick-start)
-   [Usage](#usage)
    -   [GUI Mode (Beginner-friendly)](#gui-mode-beginner-friendly)
    -   [CLI / Silent Mode](#cli--silent-mode)
    -   [Parameters](#parameters)
    -   [Command-Line Examples](#command-line-examples)
-   [Scheduling (Automatic Backups)](#scheduling-automatic-backups)
-   [Configuration](#configuration)
-   [How It Works](#how-it-works)
    -   [Profile Discovery](#profile-discovery)
    -   [Network Path Detection &
        Fallback](#network-path-detection--fallback)
    -   [Safety & Integrity Checks](#safety--integrity-checks)
    -   [Logging](#logging)
-   [File Locations](#file-locations)
-   [Troubleshooting](#troubleshooting)
-   [FAQ](#faq)
-   [Security Notes](#security-notes)
-   [Known Limitations](#known-limitations)
-   [Uninstall / Cleanup](#uninstall--cleanup)
-   [Changelog](#changelog)
-   [License](#license)
-   [Credits](#credits)

------------------------------------------------------------------------

## Overview

**Bookmark Backup Tool v5.0** is a PowerShell tool that lets you export
and import bookmarks for **Google Chrome**, **Microsoft Edge**, and
**Mozilla Firefox**. It supports **GUI** for interactive use and **CLI**
for automation, offers **scheduled backups**, **multi-profile
handling**, **file integrity checks**, and **comprehensive logging**.
It's designed to be robust in enterprise environments (network paths,
permissions, and process safety) while staying simple for individual
users.

------------------------------------------------------------------------

## Key Features

-   Multi-browser support (Chrome, Edge, Firefox)
-   Two interfaces: GUI (Windows Forms) and CLI (silent mode)
-   Automatic path detection with network share preference and Desktop
    fallback
-   Multi-profile support (latest or all profiles)
-   Automatic backup before imports
-   File integrity verification for JSON and SQLite files
-   Scheduling via Windows Task Scheduler (Daily/Weekly/Monthly)
-   Configuration file support with sensible defaults
-   Verbose logging with retention policy
-   Safety mechanisms: browser process detection, retry logic,
    WhatIf/Force switches

------------------------------------------------------------------------

## Requirements

-   Windows 10 or 11
-   PowerShell 5.1+
-   .NET Assemblies: `System.Windows.Forms`, `System.Drawing`
-   File system access to profiles/target paths
-   Admin rights for scheduled tasks (sometimes required)

------------------------------------------------------------------------

## Supported Browsers & Files

  --------------------------------------------------------------------------------------------------------------------
  Browser     Exported/Imported File              Native Location
  ----------- ----------------------------------- --------------------------------------------------------------------
  Chrome      `Chrome-Bookmarks.json`             `%LOCALAPPDATA%\\Google\\Chrome\\User Data\\<Profile>\\Bookmarks`

  Edge        `Edge-Bookmarks.json`               `%LOCALAPPDATA%\\Microsoft\\Edge\\User Data\\<Profile>\\Bookmarks`

  Firefox     `Firefox-places.sqlite`             `%APPDATA%\\Mozilla\\Firefox\\Profiles\\<Profile>\\places.sqlite`
  --------------------------------------------------------------------------------------------------------------------

------------------------------------------------------------------------

## Quick Start

``` powershell
# Launch GUI
.\BookMarkToolv5.ps1
```

------------------------------------------------------------------------

## Usage

### GUI Mode (Beginner-friendly)

Run without parameters:

``` powershell
.\BookMarkToolv5.ps1
```

### CLI / Silent Mode

Requires `-Silent` and `-Action Export|Import` with browser flags.

### Parameters

  Parameter                Description
  ------------------------ ----------------------------
  `-Silent`                Run without GUI
  `-Action`                `Export` or `Import`
  `-Chrome`                Include Chrome
  `-Edge`                  Include Edge
  `-Firefox`               Include Firefox
  `-TargetPath`            Override detected path
  `-WhatIf`                Preview only
  `-Force`                 Force without confirmation
  `-AllProfiles`           Process all profiles
  `-CreateScheduledTask`   Create auto-backup task
  `-ScheduleFrequency`     Daily/Weekly/Monthly
  `-ConfigPath`            Use custom config file

### Command-Line Examples

#### Export

``` powershell
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge -Firefox
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -TargetPath "D:\Backups"
.\BookMarkToolv5.ps1 -Silent -Action Export -Edge -Firefox -Verbose
```

#### Import

``` powershell
.\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -Edge -Firefox
.\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -TargetPath "D:\Backups"
.\BookMarkToolv5.ps1 -Silent -Action Import -Firefox -AllProfiles
```

#### WhatIf / Force

``` powershell
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -WhatIf
.\BookMarkToolv5.ps1 -Silent -Action Import -Edge -Force
```

#### Scheduling

``` powershell
.\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Daily
.\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Weekly
.\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Monthly
Unregister-ScheduledTask -TaskName "BookmarkBackupTool_AutoExport" -Confirm:$false
```

#### Configs

``` powershell
.\BookMarkToolv5.ps1 -ConfigPath "C:\MyConfigs\bookmarktool.json"
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -AllProfiles -ConfigPath "C:\MyConfigs\bookmarktool.json"
```

#### Workflows

``` powershell
# Home/Work sync
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge
.\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -Edge

# System migration
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -TargetPath "D:\Backup"
.\BookMarkToolv5.ps1 -Silent -Action Import -Edge -TargetPath "D:\Backup"

# Full backup
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge -Firefox -AllProfiles
```

------------------------------------------------------------------------

## Scheduling (Automatic Backups)

-   Daily, Weekly, Monthly options at 6:00 PM by default.\
-   Creates a Task Scheduler job named `BookmarkBackupTool_AutoExport`.

------------------------------------------------------------------------

## Configuration

Default config is stored at:

    %USERPROFILE%\\BookmarkTool.config.json

Example fields:

``` json
{
  "PreferNetworkPath": true,
  "DefaultBrowsers": ["Chrome","Edge"],
  "VerifyFileIntegrity": true,
  "AutoBackupBeforeImport": true
}
```

------------------------------------------------------------------------

## How It Works

### Profile Discovery

-   Chrome/Edge: detects `Default` and `Profile X` folders.
-   Firefox: parses `profiles.ini` and validates `places.sqlite`.

### Network Path Detection & Fallback

-   Uses `%HOMESHARE%` if valid, otherwise falls back to Desktop.

### Safety & Integrity Checks

-   Detects if browsers are running before import.
-   Auto backups created in `BookmarkTool_Backups`.
-   JSON validated for structure; SQLite validated for signature.

### Logging

-   Writes `BookmarkTool.log` in target path or user profile.
-   Retains logs for 30 days (default).

------------------------------------------------------------------------

## File Locations

-   Exports: `Chrome-Bookmarks.json`, `Edge-Bookmarks.json`,
    `Firefox-places.sqlite`
-   Logs: `%USERPROFILE%\\BookmarkTool.log` or TargetPath
-   Config: `%USERPROFILE%\\BookmarkTool.config.json`
-   Auto backups: `<Profile>\\BookmarkTool_Backups`

------------------------------------------------------------------------

## Troubleshooting

-   **Browser is running** → Close browsers or run:\
    `Get-Process chrome,msedge,firefox | Stop-Process`
-   **Network path error** → Falls back to Desktop or specify
    `-TargetPath`
-   **Permission denied** → Run PowerShell as Admin
-   **Execution policy** → Run:\
    `powershell -ExecutionPolicy Bypass -File .\BookMarkToolv5.ps1`

------------------------------------------------------------------------

## FAQ

-   **Does this merge bookmarks?** → No, it replaces files.\
-   **Can I compress backups?** → Use `Compress-Archive` after export.\
-   **Can I import/export multiple users?** → Yes, wrap in a loop across
    `C:\Users`.

------------------------------------------------------------------------

## Security Notes

-   Exported bookmarks may contain sensitive info. Secure storage is
    recommended.

------------------------------------------------------------------------

## Known Limitations

-   No merge capability.\
-   Monthly schedule = 4-week interval.\
-   Enterprise profile redirections may affect detection.

------------------------------------------------------------------------

## Uninstall / Cleanup

``` powershell
Unregister-ScheduledTask -TaskName "BookmarkBackupTool_AutoExport" -Confirm:$false
Remove-Item "$env:USERPROFILE\BookmarkTool.config.json"
Remove-Item "$env:USERPROFILE\BookmarkTool.log"
```

------------------------------------------------------------------------

## Changelog

**v5.0 -- September 22, 2025**\
- Enhanced GUI\
- Network path detection w/ retry\
- Process safety checks\
- Auto backups before import\
- File integrity validation\
- Comprehensive logging\
- Scheduling improvements

------------------------------------------------------------------------

## License

MIT License

------------------------------------------------------------------------

## Credits

Author: **Jesus M. Ayala**

------------------------------------------------------------------------
