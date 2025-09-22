# Bookmark Backup Tool v5.0 -- Enhanced Edition

PowerShell tool that lets you export and import bookmarks for Google Chrome, Microsoft Edge, and Mozilla Firefox. It supports GUI for interactive use and CLI for automation, offers scheduled backups, multi-profile handling, file integrity checks, and comprehensive logging. It’s designed to be robust in enterprise environments (network paths, permissions, and process safety) while staying simple for individual users.

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

## Command-Line Examples

Below are **detailed CLI examples** demonstrating how to use every
available option.

------------------------------------------------------------------------

### Basic Export

``` powershell
# Export bookmarks for all browsers to auto-detected path
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge -Firefox
```

``` powershell
# Export Chrome only to a specific folder
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -TargetPath "D:\Backups"
```

``` powershell
# Export Edge + Firefox with verbose output
.\BookMarkToolv5.ps1 -Silent -Action Export -Edge -Firefox -Verbose
```

------------------------------------------------------------------------

### Basic Import

``` powershell
# Import bookmarks for all browsers from auto-detected path
.\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -Edge -Firefox
```

``` powershell
# Import Chrome bookmarks from specific backup folder
.\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -TargetPath "D:\Backups"
```

``` powershell
# Import Firefox for all available profiles instead of just the most recent
.\BookMarkToolv5.ps1 -Silent -Action Import -Firefox -AllProfiles
```

------------------------------------------------------------------------

### Using `-WhatIf`

``` powershell
# Preview an export operation without actually exporting
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -WhatIf

# Preview an import from D:\Backups for Edge
.\BookMarkToolv5.ps1 -Silent -Action Import -Edge -TargetPath "D:\Backups" -WhatIf
```

------------------------------------------------------------------------

### Using `-Force`

``` powershell
# Force an import into Edge without confirmation prompts
.\BookMarkToolv5.ps1 -Silent -Action Import -Edge -Force
```

``` powershell
# Force import into Chrome from a network share
.\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -TargetPath "\\server\backups" -Force
```

------------------------------------------------------------------------

### Scheduling

``` powershell
# Create a daily scheduled export for all browsers at 6 PM
.\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Daily
```

``` powershell
# Create a weekly scheduled export (every Sunday at 6 PM)
.\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Weekly
```

``` powershell
# Create a monthly scheduled export
.\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Monthly
```

``` powershell
# Remove scheduled task manually
Unregister-ScheduledTask -TaskName "BookmarkBackupTool_AutoExport" -Confirm:$false
```

------------------------------------------------------------------------

### Configurations

``` powershell
# Run with a custom config file
.\BookMarkToolv5.ps1 -ConfigPath "C:\MyConfigsookmarktool.json"
```

``` powershell
# Export Chrome with all profiles using custom config
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -AllProfiles -ConfigPath "C:\MyConfigsookmarktool.json"
```

------------------------------------------------------------------------

### Real-World Workflows

**Home/Work Sync**

``` powershell
# At work
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge

# At home
.\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -Edge
```

**System Migration**

``` powershell
# Export Chrome bookmarks from old PC
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -TargetPath "D:\Backup"

# Import into Edge on new PC
.\BookMarkToolv5.ps1 -Silent -Action Import -Edge -TargetPath "D:\Backup"
```

**Full System Backup**

``` powershell
.\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge -Firefox -AllProfiles
```

------------------------------------------------------------------------

## Notes

-   Always close browsers before **Import**.
-   Use `-WhatIf` for safe previews.
-   Combine `-Force` with network paths in automation scenarios.
-   Scheduling replaces any existing scheduled task of the same name.

------------------------------------------------------------------------
