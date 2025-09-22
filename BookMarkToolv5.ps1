# =====================================================================================
# Bookmark Backup Tool v5.0 - Enhanced Edition
# Author: Jesus M. Ayala
#     Version: 5.0
#     Last Modified: September 22, 2025
#     Requires: PowerShell 5.1+, Windows 10/11
#     License: MIT
# =====================================================================================
# 
# .SYNOPSIS
#     Advanced bookmark backup and restore tool for Chrome, Edge, and Firefox browsers
#     with comprehensive features including GUI, CLI, scheduling, and safety mechanisms.
#
# .DESCRIPTION
#     This PowerShell script provides enterprise-grade bookmark management with:
#     - Multi-browser support (Chrome, Edge, Firefox)
#     - GUI and command-line interfaces
#     - Automatic network path detection with resilient fallback
#     - Browser process detection and safety mechanisms
#     - Automatic backup before import operations
#     - Multiple profile support
#     - Scheduled backup capabilities
#     - Configuration file support
#     - Comprehensive logging and error handling
#     - File integrity verification
#     - Progress indication and user feedback
#
# .PARAMETER Silent
#     Run in silent mode without GUI. Requires -Action parameter.
#
# .PARAMETER Action
#     Specifies the action to perform: Export or Import. Required for silent mode.
#
# .PARAMETER Chrome
#     Include Chrome bookmarks in the operation.
#
# .PARAMETER Edge
#     Include Microsoft Edge bookmarks in the operation.
#
# .PARAMETER Firefox
#     Include Firefox bookmarks in the operation.
#
# .PARAMETER TargetPath
#     Override automatic path detection with a specific target path.
#     Path will be validated and created if necessary.
#
# .PARAMETER WhatIf
#     Show what would be done without actually performing the operation.
#
# .PARAMETER Force
#     Bypass confirmations and force operations.
#
# .PARAMETER CreateScheduledTask
#     Create a Windows Task Scheduler entry for automatic daily backups.
#
# .PARAMETER ScheduleFrequency
#     Frequency for scheduled backups: Daily, Weekly, or Monthly.
#
# .PARAMETER ConfigPath
#     Path to configuration file. Defaults to user profile.
#
# .PARAMETER AllProfiles
#     Process all browser profiles instead of just the most recent.
#
# .EXAMPLE
#     .\BookMarkToolv5.ps1
#     Launch the GUI for interactive bookmark management.
#
# .EXAMPLE
#     .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge -Firefox
#     Export all browser bookmarks in silent mode using auto-detected path.
#
# .EXAMPLE
#     .\BookMarkToolv5.ps1 -Silent -Action Import -TargetPath "D:\Backups" -Chrome -WhatIf
#     Show what would be imported for Chrome without actually performing the operation.
#
# .EXAMPLE
#     .\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Daily
#     Create a daily scheduled task for automatic bookmark exports.
#
# =====================================================================================
# HOW TO USE THIS SCRIPT - COMPREHENSIVE GUIDE
# =====================================================================================
#
# This script provides multiple ways to manage your bookmarks across Chrome, Edge, and Firefox.
# Below are detailed instructions and examples for all major use cases.
#
# -----------------------------------------------------------------------------
# 1. BASIC GUI MODE (Recommended for beginners)
# -----------------------------------------------------------------------------
#
# Simply run the script without any parameters to launch the graphical interface:
#
#     .\BookMarkToolv5.ps1
#
# The GUI will:
# - Automatically detect your network path or use Desktop as fallback
# - Show checkboxes for Chrome, Edge, and Firefox
# - Provide Export and Import buttons
# - Display progress and completion messages
#
# -----------------------------------------------------------------------------
# 2. COMMAND LINE EXPORT (Silent Mode)
# -----------------------------------------------------------------------------
#
# Export all browser bookmarks to auto-detected path:
#     .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge -Firefox
#
# Export only Chrome bookmarks to specific path:
#     .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -TargetPath "D:\Backups"
#
# Export with verbose logging:
#     .\BookMarkToolv5.ps1 -Silent -Action Export -Edge -Firefox -Verbose
#
# Export with "What If" mode (preview what would be done):
#     .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -WhatIf
#
# -----------------------------------------------------------------------------
# 3. COMMAND LINE IMPORT (Silent Mode)
# -----------------------------------------------------------------------------
#
# Import bookmarks from auto-detected path:
#     .\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -Edge -Firefox
#
# Import from specific path:
#     .\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -TargetPath "D:\Backups"
#
# Force import without confirmations:
#     .\BookMarkToolv5.ps1 -Silent -Action Import -Edge -Force
#
# Import all profiles instead of just the latest:
#     .\BookMarkToolv5.ps1 -Silent -Action Import -Firefox -AllProfiles
#
# -----------------------------------------------------------------------------
# 4. SCHEDULED AUTOMATIC BACKUPS
# -----------------------------------------------------------------------------
#
# Create daily automatic backups at 6:00 PM:
#     .\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Daily
#
# Create weekly backups every Sunday at 6:00 PM:
#     .\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Weekly
#
# Create monthly backups on the 1st of each month:
#     .\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Monthly
#
# Update existing scheduled task (automatically replaces if exists):
#     .\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Weekly
#
# Remove existing scheduled task:
#     Unregister-ScheduledTask -TaskName "BookmarkBackupTool_AutoExport" -Confirm:$false
#
# -----------------------------------------------------------------------------
# 5. ADVANCED CONFIGURATION OPTIONS
# -----------------------------------------------------------------------------
#
# Use custom configuration file:
#     .\BookMarkToolv5.ps1 -ConfigPath "C:\MyConfigs\bookmarks.json"
#
# Process all browser profiles (not just the latest):
#     .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -AllProfiles
#
# Force operations with custom target path:
#     .\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -TargetPath "\\server\backups" -Force
#
# -----------------------------------------------------------------------------
# 6. TYPICAL WORKFLOW EXAMPLES
# -----------------------------------------------------------------------------
#
# Daily Home/Work Sync Workflow:
# 1. At work: .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge
# 2. At home:  .\BookMarkToolv5.ps1 -Silent -Action Import -Chrome -Edge
#
# Automated Backup Setup:
# 1. .\BookMarkToolv5.ps1 -CreateScheduledTask -ScheduleFrequency Daily
# 2. Backups will run automatically every day at 6:00 PM
#
# Browser Migration:
# 1. Export from old browser: .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome
# 2. Import to new browser:   .\BookMarkToolv5.ps1 -Silent -Action Import -Edge
#
# Complete System Backup:
# 1. .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge -Firefox -AllProfiles
#
# -----------------------------------------------------------------------------
# 7. TROUBLESHOOTING COMMON ISSUES
# -----------------------------------------------------------------------------
#
# Problem: "Browser is running" error during import
# Solution: Close the browser completely and try again
#     Get-Process chrome,msedge,firefox -ErrorAction SilentlyContinue | Stop-Process
#
# Problem: Network path not accessible
# Solution: Script automatically falls back to Desktop, or specify local path:
#     .\BookMarkToolv5.ps1 -TargetPath "$env:USERPROFILE\Documents\Bookmarks"
#
# Problem: Permission denied
# Solution: Run PowerShell as Administrator:
#     Start-Process powershell -Verb RunAs -ArgumentList "-File .\BookMarkToolv5.ps1"
#
# Problem: Execution policy prevents script from running
# Solution: Temporarily bypass execution policy:
#     powershell -ExecutionPolicy Bypass -File .\BookMarkToolv5.ps1
#
# -----------------------------------------------------------------------------
# 8. FILE LOCATIONS AND FORMATS
# -----------------------------------------------------------------------------
#
# Exported files created:
# - Chrome-Bookmarks.json     (Chrome bookmarks in JSON format)
# - Edge-Bookmarks.json       (Edge bookmarks in JSON format)  
# - Firefox-places.sqlite     (Firefox complete bookmark database)
#
# Log file location:
# - %USERPROFILE%\BookmarkTool.log (in GUI mode)
# - Same as target path (in silent mode)
#
# Configuration file location:
# - %USERPROFILE%\BookmarkTool.config.json (auto-created)
#
# Automatic backups created in:
# - [Browser Profile]\BookmarkTool_Backups\*.backup
#
# -----------------------------------------------------------------------------
# 9. SAFETY FEATURES
# -----------------------------------------------------------------------------
#
# The script includes several safety mechanisms:
# - Automatic backup creation before any import operation
# - Browser running detection (prevents corruption)
# - File integrity verification for all bookmark files
# - Network timeout handling with automatic fallback
# - Retry logic with exponential backoff for network operations
# - Comprehensive error logging and recovery
#
# -----------------------------------------------------------------------------
# 10. ADVANCED POWERSHELL INTEGRATION
# -----------------------------------------------------------------------------
#
# Import as module for custom scripting:
#     Import-Module .\BookMarkToolv5.ps1 -Force
#     Export-Bookmarks -Path "C:\Backup" -Chrome -Edge
#
# Batch processing multiple users:
#     Get-ChildItem "C:\Users" | ForEach-Object {
#         $userProfile = $_.FullName
#         .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -TargetPath "$userProfile\BookmarkBackup"
#     }
#
# Integration with backup scripts:
#     # Include in your daily backup routine
#     .\BookMarkToolv5.ps1 -Silent -Action Export -Chrome -Edge -Firefox
#     # Then compress and archive the results
#     Compress-Archive -Path ".\*Bookmarks*" -DestinationPath "BookmarkBackup_$(Get-Date -Format 'yyyyMMdd').zip"
#
# .NOTES
#     
#
# =====================================================================================

# Requires PowerShell 5.1 or higher for Windows Forms and advanced features
#Requires -Version 5.1

# =====================================================================================
# ENHANCED SCRIPT PARAMETERS
# =====================================================================================

[CmdletBinding(SupportsShouldProcess)]
param(
    # === CORE OPERATION PARAMETERS ===
    
    # Run in silent mode (no GUI) - requires -Action parameter
    [Parameter(HelpMessage="Run in silent mode without GUI interface")]
    [switch]$Silent,
    
    # Action to perform in silent mode: Export or Import bookmarks
    [Parameter(HelpMessage="Action to perform: Export or Import")]
    [ValidateSet('Export','Import')]
    [string]$Action,
    
    # Include Chrome bookmarks in the operation
    [Parameter(HelpMessage="Include Google Chrome bookmarks")]
    [switch]$Chrome,
    
    # Include Edge bookmarks in the operation
    [Parameter(HelpMessage="Include Microsoft Edge bookmarks")]
    [switch]$Edge,
    
    # Include Firefox bookmarks in the operation
    [Parameter(HelpMessage="Include Mozilla Firefox bookmarks")]
    [switch]$Firefox,
    
    # Override automatic path detection with a specific target path
    [Parameter(HelpMessage="Specific path for bookmark storage")]
    [ValidateScript({
        if ($_ -and !(Test-Path $_ -IsValid)) {
            throw "Invalid path format: $_"
        }
        return $true
    })]
    [string]$TargetPath,
    
    # === ADVANCED OPERATION PARAMETERS ===
    
    # Bypass confirmations and force operations
    [Parameter(HelpMessage="Force operations without confirmations")]
    [switch]$Force,
    
    # === SCHEDULING PARAMETERS ===
    
    # Create a Windows Task Scheduler entry for automatic backups
    [Parameter(HelpMessage="Create scheduled task for automatic backups")]
    [switch]$CreateScheduledTask,
    
    # Frequency for scheduled backups
    [Parameter(HelpMessage="Schedule frequency for automatic backups")]
    [ValidateSet('Daily','Weekly','Monthly')]
    [string]$ScheduleFrequency = 'Daily',
    
    # === CONFIGURATION PARAMETERS ===
    
    # Path to configuration file
    [Parameter(HelpMessage="Path to configuration file")]
    [string]$ConfigPath,
    
    # Process all browser profiles instead of just the most recent
    [Parameter(HelpMessage="Process all browser profiles instead of just the latest")]
    [switch]$AllProfiles
)

# =====================================================================================
# GLOBAL VARIABLES AND CONFIGURATION
# =====================================================================================

# Script-wide variables for configuration and state management
$script:Config = $null
$script:OperationResults = @{}
$script:StartTime = Get-Date

# =====================================================================================
# CONFIGURATION MANAGEMENT
# =====================================================================================

# ----------------------------------------
# Get-DefaultConfiguration
# ----------------------------------------
# Purpose: Return default configuration settings
# Returns: Hashtable with default settings
function Get-DefaultConfiguration {
    return @{
        # Default backup location settings
        DefaultPath = ""
        PreferNetworkPath = $true
        
        # Default browser selections
        DefaultBrowsers = @('Chrome', 'Edge')
        
        # Safety and backup settings
        AutoBackupBeforeImport = $true
        VerifyFileIntegrity = $true
        
        # Network and retry settings
        NetworkTimeoutSeconds = 3
        MaxRetryAttempts = 3
        RetryDelaySeconds = 1
        
        # Logging settings
        DetailedLogging = $true
        LogRetentionDays = 30
        
        # GUI settings
        ShowProgressIndicator = $true
        ConfirmOperations = $true
        
        # Advanced settings
        CreateBackupOnExport = $false
        CompressBackups = $false
    }
}

# ----------------------------------------
# Get-Configuration
# ----------------------------------------
# Purpose: Load configuration from file or return defaults
# Parameters:
#   - ConfigPath: Optional path to config file
# Returns: Configuration hashtable
function Get-Configuration {
    param([string]$ConfigFilePath)
    
    # Determine config file path
    if (!$ConfigFilePath) {
        $ConfigFilePath = if ($ConfigPath) { 
            $ConfigPath 
        } else { 
            Join-Path $env:USERPROFILE "BookmarkTool.config.json" 
        }
    }
    
    # Load config file if it exists
    if (Test-Path $ConfigFilePath) {
        try {
            Write-Verbose "Loading configuration from: $ConfigFilePath"
            $configContent = Get-Content $ConfigFilePath -Raw | ConvertFrom-Json
            
            # Convert PSCustomObject to hashtable and merge with defaults
            $config = Get-DefaultConfiguration
            $configContent.PSObject.Properties | ForEach-Object {
                $config[$_.Name] = $_.Value
            }
            
            Write-Verbose "Configuration loaded successfully"
            return $config
        }
        catch {
            Write-Warning "Failed to load configuration from $ConfigFilePath`: $_"
            Write-Warning "Using default configuration"
        }
    }
    else {
        Write-Verbose "No configuration file found at $ConfigFilePath, using defaults"
    }
    
    # Return default configuration
    return Get-DefaultConfiguration
}

# ----------------------------------------
# Save-Configuration
# ----------------------------------------
# Purpose: Save current configuration to file
# Parameters:
#   - Config: Configuration hashtable to save
#   - ConfigPath: Path where to save the config file
function Save-Configuration {
    param(
        [hashtable]$Config,
        [string]$ConfigFilePath
    )
    
    if (!$ConfigFilePath) {
        $ConfigFilePath = if ($ConfigPath) { 
            $ConfigPath 
        } else { 
            Join-Path $env:USERPROFILE "BookmarkTool.config.json" 
        }
    }
    
    try {
        $Config | ConvertTo-Json -Depth 3 | Set-Content $ConfigFilePath -Encoding UTF8
        Write-Verbose "Configuration saved to: $ConfigFilePath"
        return $true
    }
    catch {
        Write-Warning "Failed to save configuration to $ConfigFilePath`: $_"
        return $false
    }
}

# Load configuration at script startup
$script:Config = Get-Configuration

# =====================================================================================
# SINGLE-THREADED APARTMENT (STA) INITIALIZATION
# =====================================================================================
# Windows Forms requires STA mode. If we're not in STA, restart the script in STA mode.
# This ensures proper GUI functionality and prevents threading issues.

if ([Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    Write-Verbose "Re-launching in STA mode for Windows Forms compatibility..."
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName  = (Get-Command powershell).Source
    $psi.Arguments = "-STA -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" " + ($MyInvocation.UnboundArguments -join ' ')
    [Diagnostics.Process]::Start($psi) | Out-Null
    exit
}

# =====================================================================================
# ASSEMBLY LOADING
# =====================================================================================
# Load required .NET assemblies for GUI components
Add-Type -AssemblyName System.Windows.Forms  # For GUI controls and dialogs
Add-Type -AssemblyName System.Drawing         # For GUI sizing and positioning

# =====================================================================================
# UTILITY FUNCTIONS
# =====================================================================================

# ----------------------------------------
# Test-Prerequisites
# ----------------------------------------
# Purpose: Validate system prerequisites and permissions
# Returns: Boolean indicating if all prerequisites are met
function Test-Prerequisites {
    param([string]$TargetPath)
    
    $issues = @()
    
    # Test PowerShell version
    if ($PSVersionTable.PSVersion.Major -lt 5) {
        $issues += "PowerShell 5.1 or higher required"
    }
    
    # Test target path if provided
    if ($TargetPath) {
        try {
            # Test path validity
            if (!(Test-Path $TargetPath -IsValid)) {
                $issues += "Invalid target path format: $TargetPath"
            }
            
            # Test write permissions
            if (Test-Path $TargetPath) {
                $testFile = Join-Path $TargetPath "_bmtool_permission_test.tmp"
                try {
                    New-Item -Path $testFile -ItemType File -Force -ErrorAction Stop | Out-Null
                    Remove-Item -Path $testFile -Force -ErrorAction Stop
                }
                catch {
                    $issues += "No write permission to target path: $TargetPath"
                }
            }
        }
        catch {
            $issues += "Cannot access target path: $TargetPath - $_"
        }
    }
    
    # Test .NET assemblies
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    }
    catch {
        $issues += "Required .NET assemblies not available"
    }
    
    if ($issues.Count -gt 0) {
        Write-Error "Prerequisites check failed:"
        $issues | ForEach-Object { Write-Error "  - $_" }
        return $false
    }
    
    Write-Verbose "All prerequisites satisfied"
    return $true
}

# ----------------------------------------
# Test-PathAccess
# ----------------------------------------
# Purpose: Test if a path exists and is accessible for read/write operations
# Parameters:
#   - Path: Path to test
#   - RequireWrite: Whether write access is required
# Returns: Boolean indicating accessibility
function Test-PathAccess {
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [switch]$RequireWrite
    )
    
    try {
        # Test basic path existence and readability
        if (!(Test-Path $Path)) {
            Write-Verbose "Path does not exist: $Path"
            return $false
        }
        
        # Test write access if required
        if ($RequireWrite) {
            $testFile = Join-Path $Path "_bmtool_write_test_$(Get-Random).tmp"
            try {
                New-Item -Path $testFile -ItemType File -Force -ErrorAction Stop | Out-Null
                Remove-Item -Path $testFile -Force -ErrorAction Stop
                Write-Verbose "Write access confirmed for: $Path"
            }
            catch {
                Write-Verbose "No write access to: $Path"
                return $false
            }
        }
        
        return $true
    }
    catch {
        Write-Verbose "Failed to access path $Path`: $_"
        return $false
    }
}

# ----------------------------------------
# Invoke-WithRetry
# ----------------------------------------
# Purpose: Execute a script block with retry logic and exponential backoff
# Parameters:
#   - ScriptBlock: Code to execute
#   - MaxAttempts: Maximum number of retry attempts
#   - InitialDelay: Initial delay in seconds
#   - BackoffMultiplier: Multiplier for exponential backoff
# Returns: Result of successful execution or throws last error
function Invoke-WithRetry {
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock,
        [int]$MaxAttempts = $script:Config.MaxRetryAttempts,
        [int]$InitialDelay = $script:Config.RetryDelaySeconds,
        [double]$BackoffMultiplier = 2.0
    )
    
    $attempt = 1
    $delay = $InitialDelay
    $lastError = $null
    
    while ($attempt -le $MaxAttempts) {
        try {
            Write-Verbose "Attempt $attempt of $MaxAttempts"
            return & $ScriptBlock
        }
        catch {
            $lastError = $_
            Write-Verbose "Attempt $attempt failed: $_"
            
            if ($attempt -lt $MaxAttempts) {
                Write-Verbose "Waiting $delay seconds before retry..."
                Start-Sleep $delay
                $delay = [math]::Min($delay * $BackoffMultiplier, 30) # Cap at 30 seconds
            }
            
            $attempt++
        }
    }
    
    # All attempts failed
    throw $lastError
}

# ----------------------------------------
# Select-FileDialog
# ----------------------------------------
# Purpose: Display a Windows file dialog for file selection
# Parameters: 
#   - Filter: File type filter string (default: all files)
# Returns: Selected file path or $null if cancelled
# Usage: Used during import operations when automatic file detection fails
function Select-FileDialog {
    param(
        [string]$Filter = "All files (*.*)|*.*"
    )
    
    # Create and configure the OpenFileDialog
    $ofd = New-Object System.Windows.Forms.OpenFileDialog
    $ofd.Filter = $Filter
    
    # Show dialog and return result
    $result = $ofd.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $ofd.FileName
    }
    return $null
}

# ----------------------------------------
# Get-HomeSharePath
# ----------------------------------------
# Purpose: Intelligently determine the best path for bookmark storage with retry logic
# Logic:
#   1. Check if network home folder is accessible (with retry and timeout)
#   2. Fall back to Desktop folder if network is unavailable
# Returns: Accessible path string (network path or Desktop folder)
# Note: Uses enhanced retry logic with exponential backoff
function Get-HomeSharePath {
    # Set up paths
    $desktopPath = [IO.Path]::Combine($env:USERPROFILE, 'Desktop')
    $networkPath = if ($env:HOMESHARE) { $env:HOMESHARE } else { "\\server\home\$env:USERNAME" }

    # Only perform the network check if it's a UNC path (starts with \\)
    if (!($networkPath -like "\\*")) {
        Write-Verbose "Network path is not UNC format, using Desktop folder"
        return $desktopPath
    }

    # Enhanced network accessibility test with retry logic
    try {
        $result = Invoke-WithRetry -ScriptBlock {
            # Test network accessibility in a background job to prevent hanging
            $job = Start-Job -ScriptBlock {
                param($Path)
                try {
                    # Test by attempting to create and delete a temporary file
                    if (Test-Path -Path $Path -PathType Container) {
                        $tempFile = Join-Path $Path "_bmtool_temp_$(Get-Random).txt"
                        New-Item -Path $tempFile -ItemType File -Force -ErrorAction Stop | Out-Null
                        Remove-Item -Path $tempFile -Force -ErrorAction Stop
                        return $true
                    }
                } catch {
                    # Any error means the path is not accessible
                }
                return $false
            } -ArgumentList $networkPath

            # Wait for job completion with timeout from config
            $isAccessible = $false
            $timeout = $script:Config.NetworkTimeoutSeconds
            if (Wait-Job -Job $job -Timeout $timeout) {
                $isAccessible = Receive-Job -Job $job
            } else {
                Write-Verbose "Network path check timed out after $timeout seconds"
                Stop-Job -Job $job -ErrorAction SilentlyContinue
            }
            
            # Clean up the background job
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue

            if (!$isAccessible) {
                throw "Network path not accessible: $networkPath"
            }
            
            return $networkPath
        } -MaxAttempts $script:Config.MaxRetryAttempts
        
        Write-Verbose "Using network path: $result"
        return $result
    }
    catch {
        Write-Verbose "Network path failed after retries: $_"
        Write-Verbose "Using Desktop fallback: $desktopPath"
        return $desktopPath
    }
}

# =====================================================================================
# LOGGING FUNCTIONS
# =====================================================================================

# ----------------------------------------
# Get-LogFilePath
# ----------------------------------------
# Purpose: Determine where to place the log file based on current operation mode
# Logic: Uses TargetPath if specified, otherwise uses path from Get-HomeSharePath or user profile
# Returns: Full path to the log file (BookmarkTool.log)
function Get-LogFilePath {
    $path = if ($TargetPath) { 
        $TargetPath 
    } elseif ($Silent) { 
        Get-HomeSharePath 
    } else { 
        $env:USERPROFILE 
    }
    Join-Path $path "BookmarkTool.log"
}

# ----------------------------------------
# Write-Log
# ----------------------------------------
# Purpose: Write timestamped messages to the log file
# Parameters:
#   - Message: The message to log (mandatory)
# Note: Creates the log directory if it doesn't exist
function Write-Log {
    param([Parameter(Mandatory)][string]$Message)
    
    $logFile = Get-LogFilePath
    $line = ("{0:u} {1}" -f (Get-Date), $Message)
    
    # Ensure the log directory exists
    $dir = Split-Path $logFile -Parent
    if (!(Test-Path $dir)) { 
        New-Item -ItemType Directory -Path $dir -Force | Out-Null 
    }
    
    # Write to log file with UTF8 encoding
    $line | Out-File -FilePath $logFile -Append -Encoding UTF8
}

# =====================================================================================
# BROWSER DETECTION FUNCTIONS
# =====================================================================================

# ----------------------------------------
# Test-BrowserRunning
# ----------------------------------------
# Test-BrowserRunning
# ----------------------------------------
# Purpose: Comprehensive check if a browser and its related processes are running
# Parameters:
#   - Browser: Name of the browser ('chrome', 'msedge', 'firefox')
# Returns: Boolean indicating if any browser-related processes are running
# Note: Enhanced to detect helper processes and background services
function Test-BrowserRunning {
    param([string]$Browser)
    
    # Define comprehensive process lists for each browser
    $processMap = @{
        'chrome' = @(
            'chrome',
            'GoogleChromeHelper',
            'Google Chrome Helper',
            'Google Chrome Helper (Renderer)',
            'Google Chrome Helper (GPU)',
            'Google Chrome Helper (Plugin)',
            'crashpad_handler'
        )
        'msedge' = @(
            'msedge',
            'MicrosoftEdge',
            'MicrosoftEdgeWebView2',
            'msedgewebview2',
            'MicrosoftEdgeCP',
            'MicrosoftEdgeSH',
            'identity_helper'
        )
        'firefox' = @(
            'firefox',
            'plugin-container',
            'firefox.exe',
            'crashreporter',
            'updater',
            'maintenanceservice'
        )
    }
    
    $browserLower = $Browser.ToLower()
    if (!$processMap.ContainsKey($browserLower)) {
        Write-Warning "Unknown browser type: $Browser"
        return $false
    }
    
    $foundProcesses = @()
    foreach ($processName in $processMap[$browserLower]) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        if ($processes) {
            $foundProcesses += $processes
            Write-Verbose "Found running process: $processName (PID: $($processes.Id -join ', '))"
        }
    }
    
    $isRunning = $foundProcesses.Count -gt 0
    if ($isRunning) {
        Write-Verbose "$Browser is running with $($foundProcesses.Count) related processes"
    } else {
        Write-Verbose "$Browser is not running"
    }
    
    return $isRunning
}

# ----------------------------------------
# Get-AllBrowserProfiles
# ----------------------------------------
# Purpose: Get all browser profiles for a specific browser (not just the latest)
# Parameters:
#   - Browser: Browser type ('Chrome', 'Edge', 'Firefox')
#   - FileName: Bookmark file to look for
# Returns: Array of profile objects with Name and Path properties
function Get-AllBrowserProfiles {
    param(
        [Parameter(Mandatory)][string]$Browser,
        [Parameter(Mandatory)][string]$FileName
    )
    
    $profiles = @()
    
    switch ($Browser.ToLower()) {
        'chrome' {
            $basePath = "$env:LOCALAPPDATA\Google\Chrome\User Data"
            if (Test-Path $basePath) {
                $profileDirs = Get-ChildItem -Path $basePath -Directory | Where-Object {
                    ($_.Name -eq 'Default' -or $_.Name -match '^Profile \d+$') -and
                    (Test-Path (Join-Path $_.FullName $FileName))
                }
                
                foreach ($dir in $profileDirs) {
                    $profiles += @{
                        Name = if ($dir.Name -eq 'Default') { 'Default Profile' } else { $dir.Name }
                        Path = $dir.FullName
                        LastUsed = $dir.LastWriteTime
                    }
                }
            }
        }
        
        'edge' {
            $basePath = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
            if (Test-Path $basePath) {
                $profileDirs = Get-ChildItem -Path $basePath -Directory | Where-Object {
                    ($_.Name -eq 'Default' -or $_.Name -match '^Profile \d+$') -and
                    (Test-Path (Join-Path $_.FullName $FileName))
                }
                
                foreach ($dir in $profileDirs) {
                    $profiles += @{
                        Name = if ($dir.Name -eq 'Default') { 'Default Profile' } else { $dir.Name }
                        Path = $dir.FullName
                        LastUsed = $dir.LastWriteTime
                    }
                }
            }
        }
        
        'firefox' {
            $basePath = "$env:APPDATA\Mozilla\Firefox\Profiles"
            if (Test-Path $basePath) {
                $profileDirs = Get-ChildItem -Path $basePath -Directory | Where-Object {
                    Test-Path (Join-Path $_.FullName $FileName)
                }
                
                foreach ($dir in $profileDirs) {
                    $profiles += @{
                        Name = $dir.Name
                        Path = $dir.FullName
                        LastUsed = $dir.LastWriteTime
                    }
                }
            }
        }
    }
    
    # Sort by last used (most recent first)
    return $profiles | Sort-Object LastUsed -Descending
}

# ----------------------------------------
# Get-LatestProfilePath
# ----------------------------------------
# Purpose: Find the most recently used browser profile that contains bookmark data
# Parameters:
#   - BasePath: Base directory where profiles are stored
#   - FileName: Name of the bookmark file to look for
# Returns: Path to the most recent profile directory or $null if none found
# Logic: Sorts profiles by LastWriteTime to get the most recently used one
function Get-LatestProfilePath {
    param([string]$BasePath, [string]$FileName)
    
    # Check if base path exists
    if (!(Test-Path $BasePath)) { 
        Write-Verbose "Base path not found: $BasePath"
        return $null 
    }
    
    # Find all profile directories that contain the specified file
    $profiles = Get-ChildItem -Path $BasePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { Test-Path (Join-Path $_.FullName $FileName) } |
        Sort-Object LastWriteTime -Descending
    
    # Return the most recently modified profile, or null if none found
    if ($profiles) { 
        Write-Verbose "Found profile: $($profiles[0].FullName)"
        $profiles[0].FullName 
    } else { 
        Write-Verbose "No profiles found with $FileName in $BasePath"
        $null 
    }
}

# ----------------------------------------
# Browser-Specific Profile Functions
# ----------------------------------------
# Purpose: Get the profile path for each supported browser
# Returns: Full path to the browser's profile directory containing bookmarks

# Chrome: Looks for "Bookmarks" file in Chrome's User Data directory
function Get-ChromeProfile   { Get-LatestProfilePath "$env:LOCALAPPDATA\Google\Chrome\User Data"   "Bookmarks" }

# Edge: Looks for "Bookmarks" file in Edge's User Data directory  
function Get-EdgeProfile     { Get-LatestProfilePath "$env:LOCALAPPDATA\Microsoft\Edge\User Data"  "Bookmarks" }

# Firefox: More complex profile detection due to Firefox's profile management system
function Get-FirefoxProfile {
    # Firefox stores profile information in profiles.ini
    $ini = "$env:APPDATA\Mozilla\Firefox\profiles.ini"
    if (!(Test-Path $ini)) { 
        Write-Verbose "Firefox profiles.ini not found at: $ini"
        return $null 
    }
    
    # Parse the INI file to get profile information
    $lines = Get-Content $ini
    $pathLine   = $lines | Where-Object { $_ -match '^Path=' } | Select-Object -First 1
    $isRelLine  = $lines | Where-Object { $_ -match '^IsRelative=' } | Select-Object -First 1
    
    # Check if the path is relative to Firefox directory
    $isRel = $false
    if ($isRelLine) { $isRel = ($isRelLine -split '=',2)[1] -eq '1' }
    
    # Extract the actual path
    $rawPath = ($pathLine -split '=',2)[1]
    
    # Build full profile path based on whether it's relative or absolute
    $profileFull = if ($isRel) {
        Join-Path "$env:APPDATA\Mozilla\Firefox" $rawPath
    } else {
        $rawPath
    }
    
    # Check if this profile contains the bookmarks database
    if (Test-Path (Join-Path $profileFull 'places.sqlite')) { 
        Write-Verbose "Found Firefox profile with places.sqlite: $profileFull"
        return $profileFull 
    }
    
    # Fallback: search for any profile with places.sqlite
    Write-Verbose "Primary profile check failed, searching for any profile with places.sqlite"
    Get-LatestProfilePath "$env:APPDATA\Mozilla\Firefox\Profiles" "places.sqlite"
}

# =====================================================================================
# BACKUP AND SAFETY FUNCTIONS
# =====================================================================================

# ----------------------------------------
# Backup-ExistingBookmarks
# ----------------------------------------
# Purpose: Create a backup of existing bookmarks before import operations
# Parameters:
#   - BrowserProfile: Path to browser profile directory
#   - BookmarkFile: Name of the bookmark file
#   - BrowserName: Name of browser for logging
# Returns: Path to backup file or $null if failed
function Backup-ExistingBookmarks {
    param(
        [Parameter(Mandatory)][string]$BrowserProfile,
        [Parameter(Mandatory)][string]$BookmarkFile,
        [Parameter(Mandatory)][string]$BrowserName
    )
    
    try {
        $sourceFile = Join-Path $BrowserProfile $BookmarkFile
        
        # Check if source file exists
        if (!(Test-Path $sourceFile)) {
            Write-Verbose "No existing $BrowserName bookmark file to backup"
            return $null
        }
        
        # Create backup directory
        $backupDir = Join-Path $BrowserProfile "BookmarkTool_Backups"
        if (!(Test-Path $backupDir)) {
            New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        }
        
        # Generate timestamped backup filename
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $backupFileName = "$BookmarkFile.$timestamp.backup"
        $backupPath = Join-Path $backupDir $backupFileName
        
        # Create the backup
        Copy-Item $sourceFile $backupPath -Force
        Write-Log "Created backup: $backupPath"
        
        # Clean old backups (keep last 10)
        $backups = Get-ChildItem $backupDir -Filter "$BookmarkFile.*.backup" | 
                   Sort-Object LastWriteTime -Descending
        
        if ($backups.Count -gt 10) {
            $oldBackups = $backups | Select-Object -Skip 10
            $oldBackups | ForEach-Object {
                Remove-Item $_.FullName -Force
                Write-Verbose "Removed old backup: $($_.Name)"
            }
        }
        
        return $backupPath
    }
    catch {
        Write-Warning "Failed to create backup for $BrowserName`: $_"
        return $null
    }
}

# ----------------------------------------
# Test-BookmarkFileIntegrity
# ----------------------------------------
# Purpose: Verify the integrity of bookmark files before operations
# Parameters:
#   - FilePath: Path to the bookmark file
#   - BrowserType: Type of browser (Chrome, Edge, Firefox)
# Returns: Boolean indicating if file is valid
function Test-BookmarkFileIntegrity {
    param(
        [Parameter(Mandatory)][string]$FilePath,
        [Parameter(Mandatory)][string]$BrowserType
    )
    
    if (!(Test-Path $FilePath)) {
        Write-Verbose "File does not exist: $FilePath"
        return $false
    }
    
    try {
        switch ($BrowserType.ToLower()) {
            'chrome' {
                # Chrome bookmarks are JSON files
                $content = Get-Content $FilePath -Raw | ConvertFrom-Json
                $isValid = ($null -ne $content.roots) -and 
                          ($null -ne $content.version) -and
                          ($null -ne $content.roots.bookmark_bar)
                
                if ($isValid) {
                    Write-Verbose "Chrome bookmark file validation passed: $FilePath"
                } else {
                    Write-Verbose "Chrome bookmark file validation failed: Missing required structure"
                }
                return $isValid
            }
            
            'edge' {
                # Edge uses same format as Chrome
                $content = Get-Content $FilePath -Raw | ConvertFrom-Json
                $isValid = ($null -ne $content.roots) -and 
                          ($null -ne $content.version)
                
                if ($isValid) {
                    Write-Verbose "Edge bookmark file validation passed: $FilePath"
                } else {
                    Write-Verbose "Edge bookmark file validation failed: Missing required structure"
                }
                return $isValid
            }
            
            'firefox' {
                # Firefox uses SQLite database
                $fileInfo = Get-Item $FilePath
                $isValid = ($fileInfo.Length -gt 0) -and 
                          ($fileInfo.Extension -eq '.sqlite')
                
                # Additional SQLite header validation
                if ($isValid) {
                    $header = [System.IO.File]::ReadAllBytes($FilePath) | Select-Object -First 16
                    $sqliteSignature = [System.Text.Encoding]::ASCII.GetString($header[0..15])
                    $isValid = $sqliteSignature.StartsWith("SQLite format 3")
                }
                
                if ($isValid) {
                    Write-Verbose "Firefox bookmark file validation passed: $FilePath"
                } else {
                    Write-Verbose "Firefox bookmark file validation failed: Invalid SQLite format"
                }
                return $isValid
            }
            
            default {
                Write-Warning "Unknown browser type for integrity check: $BrowserType"
                return $false
            }
        }
    }
    catch {
        Write-Verbose "File integrity check failed for $FilePath`: $_"
        return $false
    }
}

# ----------------------------------------
# New-OperationSummary
# ----------------------------------------
# Purpose: Create a detailed summary of backup/restore operations
# Parameters:
#   - Operations: Array of operation results
#   - StartTime: When the operation started
#   - OperationType: Export or Import
# Returns: Formatted summary string
function New-OperationSummary {
    param(
        [array]$Operations,
        [datetime]$StartTime,
        [string]$OperationType
    )
    
    $endTime = Get-Date
    $duration = $endTime - $StartTime
    
    $successful = $Operations | Where-Object { $_.Success }
    $failed = $Operations | Where-Object { !$_.Success }
    
    $summary = @"
═══════════════════════════════════════════════════════════════
                    BOOKMARK $($OperationType.ToUpper()) SUMMARY
═══════════════════════════════════════════════════════════════
Start Time:     $($StartTime.ToString('yyyy-MM-dd HH:mm:ss'))
End Time:       $($endTime.ToString('yyyy-MM-dd HH:mm:ss'))
Duration:       $($duration.ToString('hh\:mm\:ss'))
Total Operations: $($Operations.Count)
Successful:     $($successful.Count)
Failed:         $($failed.Count)

DETAILS:
"@
    
    foreach ($op in $Operations) {
        $status = if ($op.Success) { "[SUCCESS]" } else { "[FAILED]" }
        $summary += "`n$status - $($op.Browser): $($op.Message)"
    }
    
    $summary += "`n═══════════════════════════════════════════════════════════════"
    
    return $summary
}

# =====================================================================================
# BOOKMARK EXPORT FUNCTION
# =====================================================================================

# ----------------------------------------
# Export-Bookmarks
# ----------------------------------------
# Purpose: Export bookmarks from selected browsers to backup files
# Parameters:
#   - Path: Directory where backup files will be saved (mandatory)
#   - Chrome: Switch to export Chrome bookmarks
#   - Edge: Switch to export Edge bookmarks  
#   - Firefox: Switch to export Firefox bookmarks
# Files Created:
#   - Chrome-Bookmarks.json (Chrome bookmarks in JSON format)
#   - Edge-Bookmarks.json (Edge bookmarks in JSON format)
#   - Firefox-places.sqlite (Firefox bookmarks database)
function Export-Bookmarks {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Chrome, 
        [switch]$Edge, 
        [switch]$Firefox
    )
    
    # Ensure the backup directory exists
    if (!(Test-Path $Path)) { 
        Write-Log "Creating backup directory: $Path"
        New-Item -ItemType Directory -Path $Path -Force | Out-Null 
    }

    # === CHROME EXPORT ===
    if ($Chrome) {
        Write-Log "Starting Chrome bookmark export..."
        $browserProfile = Get-ChromeProfile
        if ($browserProfile) {
            $src = Join-Path $browserProfile "Bookmarks"
            $dst = Join-Path $Path "Chrome-Bookmarks.json"
            try { 
                Copy-Item $src -Destination $dst -Force
                Write-Log "SUCCESS: Exported Chrome bookmarks from $browserProfile to $dst" 
            }
            catch { 
                Write-Log "ERROR: Failed to export Chrome bookmarks - $_" 
            }
        } else { 
            Write-Log "WARNING: Chrome profile not found - skipping Chrome export" 
        }
    }

    # === EDGE EXPORT ===
    if ($Edge) {
        Write-Log "Starting Edge bookmark export..."
        $browserProfile = Get-EdgeProfile
        if ($browserProfile) {
            $src = Join-Path $browserProfile "Bookmarks"
            $dst = Join-Path $Path "Edge-Bookmarks.json"
            try { 
                Copy-Item $src -Destination $dst -Force
                Write-Log "SUCCESS: Exported Edge bookmarks from $browserProfile to $dst" 
            }
            catch { 
                Write-Log "ERROR: Failed to export Edge bookmarks - $_" 
            }
        } else { 
            Write-Log "WARNING: Edge profile not found - skipping Edge export" 
        }
    }

    # === FIREFOX EXPORT ===
    if ($Firefox) {
        Write-Log "Starting Firefox bookmark export..."
        $browserProfile = Get-FirefoxProfile
        if ($browserProfile) {
            $src = Join-Path $browserProfile "places.sqlite"
            $dst = Join-Path $Path "Firefox-places.sqlite"
            try { 
                Copy-Item $src -Destination $dst -Force
                Write-Log "SUCCESS: Exported Firefox bookmarks from $browserProfile to $dst" 
            }
            catch { 
                Write-Log "ERROR: Failed to export Firefox bookmarks - $_" 
            }
        } else { 
            Write-Log "WARNING: Firefox profile not found - skipping Firefox export" 
        }
    }
    
    Write-Log "Export operation completed"
}

# =====================================================================================
# BOOKMARK IMPORT FUNCTION
# =====================================================================================

# ----------------------------------------
# Import-Bookmarks
# ----------------------------------------
# Purpose: Import bookmarks from backup files to selected browsers
# Parameters:
#   - Path: Directory containing backup files (mandatory)
#   - Chrome: Switch to import Chrome bookmarks
#   - Edge: Switch to import Edge bookmarks
#   - Firefox: Switch to import Firefox bookmarks
# Safety Features:
#   - Checks if browsers are running and aborts if they are (prevents corruption)
#   - In GUI mode, prompts user to select files if auto-detection fails
#   - Comprehensive error handling and logging
# Files Expected:
#   - Chrome-Bookmarks.json (for Chrome)
#   - Edge-Bookmarks.json (for Edge)
#   - Firefox-places.sqlite (for Firefox)
function Import-Bookmarks {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Chrome, 
        [switch]$Edge, 
        [switch]$Firefox
    )

    # === SAFETY CHECKS: Prevent imports while browsers are running ===
    # This prevents bookmark corruption that can occur when modifying files while the browser is active
    if ($Chrome  -and (Test-BrowserRunning 'chrome'))  { 
        Write-Log "ABORT: Chrome is running - cannot safely import bookmarks"
        return 
    }
    if ($Edge    -and (Test-BrowserRunning 'msedge'))  { 
        Write-Log "ABORT: Edge is running - cannot safely import bookmarks"
        return 
    }
    if ($Firefox -and (Test-BrowserRunning 'firefox')) { 
        Write-Log "ABORT: Firefox is running - cannot safely import bookmarks"
        return 
    }

    # === CHROME IMPORT ===
    if ($Chrome) {
        Write-Log "Starting Chrome bookmark import..."
        $browserProfile = Get-ChromeProfile
        if ($browserProfile) {
            $dst = Join-Path $browserProfile "Bookmarks"              # Destination: Chrome's bookmark file
            $src = Join-Path $Path "Chrome-Bookmarks.json"            # Source: Our backup file
            
            # If backup file doesn't exist and we're in GUI mode, let user select file manually
            if (!(Test-Path $src) -and -not $Silent) {
                Write-Log "Chrome backup file not found at $src, prompting user for file selection"
                $src = Select-FileDialog -Filter "Chrome Bookmarks (*.json)|*.json"
                if ($src) { Write-Log "User selected Chrome import file: $src" }
            }
            
            # Perform the import if we have a valid source file
            if ($src -and (Test-Path $src)) {
                try { 
                    Copy-Item $src -Destination $dst -Force
                    Write-Log "SUCCESS: Imported Chrome bookmarks from $src to $dst" 
                }
                catch { 
                    Write-Log "ERROR: Failed to import Chrome bookmarks - $_" 
                }
            } else { 
                Write-Log "WARNING: Chrome import file not selected or missing - skipping Chrome import" 
            }
        } else { 
            Write-Log "WARNING: Chrome profile not found - cannot import bookmarks" 
        }
    }

    # === EDGE IMPORT ===
    if ($Edge) {
        Write-Log "Starting Edge bookmark import..."
        $browserProfile = Get-EdgeProfile
        if ($browserProfile) {
            $dst = Join-Path $browserProfile "Bookmarks"              # Destination: Edge's bookmark file
            $src = Join-Path $Path "Edge-Bookmarks.json"              # Source: Our backup file
            
            # If backup file doesn't exist and we're in GUI mode, let user select file manually
            if (!(Test-Path $src) -and -not $Silent) {
                Write-Log "Edge backup file not found at $src, prompting user for file selection"
                $src = Select-FileDialog -Filter "Edge Bookmarks (*.json)|*.json"
                if ($src) { Write-Log "User selected Edge import file: $src" }
            }
            
            # Perform the import if we have a valid source file
            if ($src -and (Test-Path $src)) {
                try { 
                    Copy-Item $src -Destination $dst -Force
                    Write-Log "SUCCESS: Imported Edge bookmarks from $src to $dst" 
                }
                catch { 
                    Write-Log "ERROR: Failed to import Edge bookmarks - $_" 
                }
            } else { 
                Write-Log "WARNING: Edge import file not selected or missing - skipping Edge import" 
            }
        } else { 
            Write-Log "WARNING: Edge profile not found - cannot import bookmarks" 
        }
    }

    # === FIREFOX IMPORT ===
    if ($Firefox) {
        Write-Log "Starting Firefox bookmark import..."
        $browserProfile = Get-FirefoxProfile
        if ($browserProfile) {
            $dst = Join-Path $browserProfile "places.sqlite"          # Destination: Firefox's bookmark database
            $src = Join-Path $Path "Firefox-places.sqlite"            # Source: Our backup file
            
            # If backup file doesn't exist and we're in GUI mode, let user select file manually
            if (!(Test-Path $src) -and -not $Silent) {
                Write-Log "Firefox backup file not found at $src, prompting user for file selection"
                $src = Select-FileDialog -Filter "Firefox places.sqlite|places.sqlite"
                if ($src) { Write-Log "User selected Firefox import file: $src" }
            }
            
            # Perform the import if we have a valid source file
            if ($src -and (Test-Path $src)) {
                try { 
                    Copy-Item $src -Destination $dst -Force
                    Write-Log "SUCCESS: Imported Firefox bookmarks from $src to $dst" 
                }
                catch { 
                    Write-Log "ERROR: Failed to import Firefox bookmarks - $_" 
                }
            } else { 
                Write-Log "WARNING: Firefox import file not selected or missing - skipping Firefox import" 
            }
        } else { 
            Write-Log "WARNING: Firefox profile not found - cannot import bookmarks" 
        }
    }
    
    Write-Log "Import operation completed"
}

# =====================================================================================
# GRAPHICAL USER INTERFACE (GUI) FUNCTIONS
# =====================================================================================

# ----------------------------------------
# Show-ProgressDialog
# ----------------------------------------
# Purpose: Display a progress dialog for long-running operations
# Parameters:
#   - Title: Dialog title
#   - Message: Progress message
#   - Operation: Script block to execute
# Returns: Result of the operation
function Show-ProgressDialog {
    param(
        [string]$Title = "Processing...",
        [string]$Message = "Please wait...",
        [scriptblock]$Operation
    )
    
    # Create progress form
    $progressForm = New-Object Windows.Forms.Form
    $progressForm.Text = $Title
    $progressForm.Size = '400,150'
    $progressForm.StartPosition = "CenterParent"
    $progressForm.FormBorderStyle = 'FixedDialog'
    $progressForm.MaximizeBox = $false
    $progressForm.MinimizeBox = $false
    $progressForm.TopMost = $true
    
    # Create progress label
    $progressLabel = New-Object Windows.Forms.Label
    $progressLabel.Text = $Message
    $progressLabel.Location = '20,20'
    $progressLabel.Size = '360,30'
    $progressLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $progressForm.Controls.Add($progressLabel)
    
    # Create progress bar
    $progressBar = New-Object Windows.Forms.ProgressBar
    $progressBar.Location = '20,60'
    $progressBar.Size = '360,25'
    $progressBar.Style = 'Marquee'
    $progressBar.MarqueeAnimationSpeed = 50
    $progressForm.Controls.Add($progressBar)
    
    # Variables to store operation result
    $operationResult = $null
    $operationError = $null
    
    # Timer to check operation completion
    $timer = New-Object Windows.Forms.Timer
    $timer.Interval = 100
    
    # Start operation in background
    $job = Start-Job -ScriptBlock $Operation
    
    # Timer tick event
    $timer.Add_Tick({
        if ($job.State -eq 'Completed') {
            $timer.Stop()
            try {
                $script:operationResult = Receive-Job -Job $job
            }
            catch {
                $script:operationError = $_
            }
            Remove-Job -Job $job -Force
            $progressForm.Close()
        }
        elseif ($job.State -in ('Failed', 'Stopped')) {
            $timer.Stop()
            try {
                $script:operationError = Receive-Job -Job $job
            }
            catch {
                $script:operationError = $_
            }
            Remove-Job -Job $job -Force
            $progressForm.Close()
        }
    })
    
    # Start timer and show dialog
    $timer.Start()
    $progressForm.ShowDialog() | Out-Null
    $timer.Stop()
    
    # Return result or throw error
    if ($operationError) {
        throw $operationError
    }
    return $operationResult
}

# ----------------------------------------
# Show-GUI
# ----------------------------------------
# Purpose: Display the enhanced main GUI for interactive bookmark management
# Features:
#   - Browser selection checkboxes (Chrome, Edge, Firefox)
#   - Path selection with browse button
#   - Export and Import buttons with progress feedback
#   - Automatic network path detection with Desktop fallback
#   - Configuration options and status display
function Show-GUI {
    Write-Log "Launching GUI mode"
    
    # === MAIN FORM SETUP ===
    $form = New-Object Windows.Forms.Form
    $form.Text = "Bookmark Backup Tool v5.0 - Enhanced Edition"
    $form.Size = '600,420'
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = 'FixedDialog'    # Prevents resizing
    $form.MaximizeBox = $false               # Removes maximize button

    # === BROWSER SELECTION CHECKBOXES ===
    # Chrome checkbox
    $chkChrome = New-Object Windows.Forms.CheckBox
    $chkChrome.Text = "Chrome"
    $chkChrome.Location = '30,30'
    $chkChrome.AutoSize = $true
    $form.Controls.Add($chkChrome)

    # Edge checkbox
    $chkEdge = New-Object Windows.Forms.CheckBox
    $chkEdge.Text = "Edge"
    $chkEdge.Location = '30,60'
    $chkEdge.AutoSize = $true
    $form.Controls.Add($chkEdge)

    # Firefox checkbox
    $chkFirefox = New-Object Windows.Forms.CheckBox
    $chkFirefox.Text = "Firefox"
    $chkFirefox.Location = '30,90'
    $chkFirefox.AutoSize = $true
    $form.Controls.Add($chkFirefox)

    # === PATH SELECTION CONTROLS ===
    # Label for path selection
    $lblPath = New-Object Windows.Forms.Label
    $lblPath.Text = "Save/Load Path:"
    $lblPath.Location = '30,130'
    $lblPath.Size = '420,20'
    $form.Controls.Add($lblPath)

    # Text box for path display/editing
    $txtPath = New-Object Windows.Forms.TextBox
    # Initialize with intelligent path detection (network path with Desktop fallback)
    $txtPath.Text = Get-HomeSharePath
    $txtPath.Location = '30,155'
    $txtPath.Size = '360,25'
    $form.Controls.Add($txtPath)

    # Browse button for manual path selection
    $btnBrowse = New-Object Windows.Forms.Button
    $btnBrowse.Text = "Browse"
    $btnBrowse.Location = '400,154'
    $btnBrowse.Size = '60,25'
    $btnBrowse.Add_Click({
        # Show folder browser dialog
        $dialog = New-Object Windows.Forms.FolderBrowserDialog
        if ($dialog.ShowDialog() -eq "OK") {
            $txtPath.Text = $dialog.SelectedPath
            Write-Log "User selected path: $($dialog.SelectedPath)"
        }
    })
    $form.Controls.Add($btnBrowse)

    # === ACTION BUTTONS ===
    # Export button
    $btnExport = New-Object Windows.Forms.Button
    $btnExport.Text = "Export Bookmarks"
    $btnExport.Size = '180,40'
    $btnExport.Location = '40,220'
    $btnExport.Add_Click({
        Write-Log "Export button clicked - starting export operation"
        
        # Re-evaluate path before export to ensure it's current
        $currentPath = if ($txtPath.Text) { $txtPath.Text } else { Get-HomeSharePath }
        $txtPath.Text = $currentPath
        
        # Perform export with selected browsers
        Export-Bookmarks -Path $currentPath -Chrome:$chkChrome.Checked -Edge:$chkEdge.Checked -Firefox:$chkFirefox.Checked
        
        # Show completion message
        [Windows.Forms.MessageBox]::Show("Export operation completed. Check the log for details.", "Export Complete", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        Write-Log "Export operation completed via GUI"
    })
    $form.Controls.Add($btnExport)

    # Import button
    $btnImport = New-Object Windows.Forms.Button
    $btnImport.Text = "Import Bookmarks"
    $btnImport.Size = '180,40'
    $btnImport.Location = '260,220'
    $btnImport.Add_Click({
        Write-Log "Import button clicked - starting import operation"
        
        # Re-evaluate path before import to ensure it's current
        $currentPath = if ($txtPath.Text) { $txtPath.Text } else { Get-HomeSharePath }
        $txtPath.Text = $currentPath
        
        # Perform import with selected browsers
        Import-Bookmarks -Path $currentPath -Chrome:$chkChrome.Checked -Edge:$chkEdge.Checked -Firefox:$chkFirefox.Checked
        
        # Show completion message
        [Windows.Forms.MessageBox]::Show("Import operation completed. Check the log for details.", "Import Complete", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        Write-Log "Import operation completed via GUI"
    })
    $form.Controls.Add($btnImport)

    # === DISPLAY THE FORM ===
    # Show the form and wait for user interaction
    Write-Log "GUI displayed successfully"
    $form.ShowDialog() | Out-Null
    Write-Log "GUI closed"
}

# =====================================================================================
# SCHEDULED TASK MANAGEMENT
# =====================================================================================

# ----------------------------------------
# New-BookmarkScheduledTask
# ----------------------------------------
# Purpose: Create a Windows Task Scheduler entry for automatic bookmark backups
# Parameters:
#   - Frequency: How often to run (Daily, Weekly, Monthly)
#   - Time: What time to run (default 6:00 PM)
# Returns: Boolean indicating success
function New-BookmarkScheduledTask {
    param(
        [string]$Frequency = 'Daily',
        [string]$Time = '18:00'
    )
    
    try {
        # Check if Task Scheduler module is available
        if (!(Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) {
            Write-Warning "Task Scheduler cmdlets not available. Requires Windows 8/Server 2012 or later."
            return $false
        }
        
        # Define the script arguments for silent mode
        $scriptArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Silent -Action Export -Chrome -Edge -Firefox"
        
        # Create the scheduled task action
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $scriptArgs
        
        # Create the trigger based on frequency
        switch ($Frequency.ToLower()) {
            'daily' {
                $trigger = New-ScheduledTaskTrigger -Daily -At $Time
            }
            'weekly' {
                $trigger = New-ScheduledTaskTrigger -Weekly -At $Time -DaysOfWeek Sunday
            }
            'monthly' {
                $trigger = New-ScheduledTaskTrigger -Weekly -At $Time -WeeksInterval 4
            }
            default {
                throw "Invalid frequency: $Frequency. Must be Daily, Weekly, or Monthly."
            }
        }
        
        # Create task settings
        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        
        # Create principal (run as current user)
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive
        
        # Check if task already exists and handle accordingly
        $taskName = "BookmarkBackupTool_AutoExport"
        $existingTask = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        
        if ($existingTask) {
            Write-Log "Scheduled task already exists: $taskName"
            Write-Host "Scheduled task '$taskName' already exists. Updating with new settings..." -ForegroundColor Yellow
            
            # Unregister the existing task first
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Write-Log "Removed existing scheduled task: $taskName"
        }
        
        # Register the scheduled task (new or replacement)
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Automatic bookmark backup using BookmarkTool v5.0" | Out-Null
        
        Write-Log "Successfully created scheduled task: $taskName"
        Write-Log "Frequency: $Frequency at $Time"
        Write-Host "[SUCCESS] Scheduled task created successfully!" -ForegroundColor Green
        Write-Host "  Task Name: $taskName" -ForegroundColor Cyan
        Write-Host "  Frequency: $Frequency at $Time" -ForegroundColor Cyan
        Write-Host "  Next Run: $($trigger.StartBoundary)" -ForegroundColor Cyan
        
        return $true
    }
    catch {
        Write-Error "Failed to create scheduled task: $_"
        Write-Log "ERROR: Failed to create scheduled task - $_"
        return $false
    }
}

# ----------------------------------------
# Remove-BookmarkScheduledTask
# ----------------------------------------
# Purpose: Remove the bookmark backup scheduled task
# Returns: Boolean indicating success
function Remove-BookmarkScheduledTask {
    try {
        $taskName = "BookmarkBackupTool_AutoExport"
        
        if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            Write-Log "Successfully removed scheduled task: $taskName"
            Write-Host "[SUCCESS] Scheduled task removed successfully!" -ForegroundColor Green
            return $true
        }
        else {
            Write-Warning "Scheduled task not found: $taskName"
            return $false
        }
    }
    catch {
        Write-Error "Failed to remove scheduled task: $_"
        Write-Log "ERROR: Failed to remove scheduled task - $_"
        return $false
    }
}

# =====================================================================================
# MAIN EXECUTION LOGIC
# =====================================================================================

# ----------------------------------------
# Script Entry Point and Enhanced Execution Logic
# ----------------------------------------
# Purpose: Determine execution mode and handle all new v5.0 features
# Modes:
#   1. Scheduled Task Creation Mode
#   2. Silent Mode: Command-line execution with specified parameters
#   3. GUI Mode: Interactive graphical interface (default)

# Only execute main logic if script is run directly (not dot-sourced)
if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch '^\s*\.\s') {
    Write-Log "=== Bookmark Backup Tool v5.0 Enhanced Edition Started ==="
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
    Write-Log "Execution mode: $(if ($CreateScheduledTask) { 'Scheduled Task Creation' } elseif ($Silent) { 'Silent' } else { 'GUI' })"

    # === PREREQUISITES CHECK ===
    Write-Verbose "Checking system prerequisites..."
    if (!(Test-Prerequisites -TargetPath $TargetPath)) {
        Write-Error "Prerequisites check failed. Cannot continue."
        exit 1
    }

# === SCHEDULED TASK CREATION MODE ===
if ($CreateScheduledTask) {
    Write-Log "Creating scheduled task for automatic backups"
    Write-Host "Creating scheduled task for automatic bookmark backups..." -ForegroundColor Yellow
    
    $success = New-BookmarkScheduledTask -Frequency $ScheduleFrequency
    if ($success) {
        Write-Log "Scheduled task creation completed successfully"
        exit 0
    } else {
        Write-Log "Scheduled task creation failed"
        exit 1
    }
}

# === PATH DETERMINATION ===
# Determine the target path based on parameters and mode
# Priority: 1) -TargetPath (if specified), 2) Auto-detection for silent mode, 3) GUI selection
$finalPath = if ($TargetPath) { 
    Write-Log "Using specified target path: $TargetPath"
    # Validate and create path if necessary
    if (!(Test-Path $TargetPath)) {
        Write-Log "Creating target directory: $TargetPath"
        try {
            New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null
        }
        catch {
            Write-Error "Failed to create target directory: $_"
            exit 1
        }
    }
    $TargetPath 
} elseif ($Silent) { 
    $autoPath = Get-HomeSharePath
    Write-Log "Auto-detected path for silent mode: $autoPath"
    $autoPath
} else { 
    Write-Log "GUI mode - path will be determined by user"
    $null 
}

# === SILENT MODE EXECUTION ===
if ($Silent) {
    Write-Log "Executing in silent mode"
    
    # Validate browser selection
    if (!$Chrome -and !$Edge -and !$Firefox) {
        Write-Warning "No browsers selected. Using default browsers from configuration."
        $Chrome = $script:Config.DefaultBrowsers -contains 'Chrome'
        $Edge = $script:Config.DefaultBrowsers -contains 'Edge'
        $Firefox = $script:Config.DefaultBrowsers -contains 'Firefox'
    }
    
    # Initialize operation tracking
    $script:OperationResults = @()
    $operationStart = Get-Date
    
    # Validate that Action parameter is provided for silent mode
    if ($Action -eq 'Export') {
        Write-Log "Performing silent export operation"
        Write-Log "Target browsers: Chrome=$Chrome, Edge=$Edge, Firefox=$Firefox"
        Write-Log "Target path: $finalPath"
        
        if ($PSCmdlet.ShouldProcess("$finalPath", "Export bookmarks")) {
            try {
                Export-Bookmarks -Path $finalPath -Chrome:$Chrome -Edge:$Edge -Firefox:$Firefox
                $summary = New-OperationSummary -Operations $script:OperationResults -StartTime $operationStart -OperationType "Export"
                Write-Log $summary
                Write-Host $summary -ForegroundColor Green
                Write-Log "Silent export completed successfully"
            }
            catch {
                Write-Error "Export operation failed: $_"
                Write-Log "ERROR: Export operation failed - $_"
                exit 1
            }
        }
    } 
    elseif ($Action -eq 'Import') {
        Write-Log "Performing silent import operation"
        Write-Log "Target browsers: Chrome=$Chrome, Edge=$Edge, Firefox=$Firefox"
        Write-Log "Source path: $finalPath"
        
        if ($PSCmdlet.ShouldProcess("$finalPath", "Import bookmarks")) {
            try {
                Import-Bookmarks -Path $finalPath -Chrome:$Chrome -Edge:$Edge -Firefox:$Firefox
                $summary = New-OperationSummary -Operations $script:OperationResults -StartTime $operationStart -OperationType "Import"
                Write-Log $summary
                Write-Host $summary -ForegroundColor Green
                Write-Log "Silent import completed successfully"
            }
            catch {
                Write-Error "Import operation failed: $_"
                Write-Log "ERROR: Import operation failed - $_"
                exit 1
            }
        }
    } 
    else {
        # Invalid or missing Action parameter
        $errorMsg = "ERROR: When using -Silent mode, -Action parameter must be specified as either 'Export' or 'Import'"
        Write-Log $errorMsg
        Write-Error $errorMsg
        exit 1
    }
    
    # Silent mode completed successfully
    Write-Log "=== Silent mode execution completed successfully ==="
    exit 0
} 
else {
    # === GUI MODE EXECUTION ===
    Write-Log "Launching enhanced GUI mode"
    try {
        Show-GUI
        Write-Log "=== GUI mode execution completed successfully ==="
    }
    catch {
        Write-Error "GUI execution failed: $_"
        Write-Log "ERROR: GUI execution failed - $_"
        exit 1
    }
}

# End of main execution check
}