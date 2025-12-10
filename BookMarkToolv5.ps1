# =====================================================================================
# Bookmark Backup Tool v5.2 - Enhanced Edition (Full Feature Parity with C# App)
# Author: Jesus M. Ayala
# Version: 5.2
# Last Modified: December 10th, 2025
# Requires: PowerShell 5.1+, Windows 10/11, .NET Framework (for System.Data.SQLite)
# License: MIT
# 
# NEW IN v5.2:
# - 🔄 Browser Auto-Close: Graceful close with 3s wait, force kill if needed
# - 📄 HTML Export/Conversion: Real Chrome JSON → HTML and Firefox SQLite → HTML
# - 📦 ZIP Archive Support: Create and extract ZIP archives of bookmarks
# - 🌐 All-Profiles Support: Export/import from all browser profiles
# - ⚠️ Interactive Browser Warnings: Y/N prompts when browsers are running
# - 🔄 Retry Logic: Auto-retry after browser closure with 2s delay
# - 📁 ZIP Import: Direct import from ZIP archives with auto-detection
# =====================================================================================

#Requires -Version 5.1

[CmdletBinding(SupportsShouldProcess)]
param(
    # === CORE OPERATION PARAMETERS ===
    [Parameter(HelpMessage="Run in silent mode without GUI interface")]
    [switch]$Silent,

    [Parameter(HelpMessage="Action to perform: Export or Import")]
    [ValidateSet('Export','Import')]
    [string]$Action,

    [Parameter(HelpMessage="Include Google Chrome bookmarks")]
    [switch]$Chrome,

    [Parameter(HelpMessage="Include Microsoft Edge bookmarks")]
    [switch]$Edge,

    [Parameter(HelpMessage="Include Mozilla Firefox bookmarks")]
    [switch]$Firefox,

    [Parameter(HelpMessage="Specific path for bookmark storage")]
    [ValidateScript({ if ($_ -and !(Test-Path $_ -IsValid)) { throw "Invalid path format: $_" }; $true })]
    [string]$TargetPath,

    # === ADVANCED OPERATION PARAMETERS ===
    [Parameter(HelpMessage="Export only HTML format (lightweight export)")]
    [switch]$HtmlOnly,

    [Parameter(HelpMessage="Process all browser profiles instead of just the latest")]
    [switch]$AllProfiles,

    [Parameter(HelpMessage="Create ZIP archive of exported bookmarks")]
    [switch]$CreateZip,

    [Parameter(HelpMessage="Force operations without confirmations")]
    [switch]$Force,

    # === SCHEDULING PARAMETERS ===
    [Parameter(HelpMessage="Create scheduled task for automatic backups")]
    [switch]$CreateScheduledTask,

    [Parameter(HelpMessage="Schedule frequency for automatic backups")]
    [ValidateSet('Daily','Weekly','Monthly')]
    [string]$ScheduleFrequency = 'Daily',

    # === CONFIGURATION PARAMETERS ===
    [Parameter(HelpMessage="Path to configuration file")]
    [string]$ConfigPath
)

# Hardening & sane defaults
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'

# =====================================================================================
# HELPER FUNCTION: Auto-Install System.Data.SQLite from NuGet
# =====================================================================================
function Install-SQLiteIfMissing {
    try {
        # Create a local lib folder if it doesn't exist
        $libPath = Join-Path $PSScriptRoot "lib"
        if (-not (Test-Path $libPath)) { New-Item -ItemType Directory -Path $libPath -Force | Out-Null }
        
        # Define paths for architecture-specific folders
        $x64Path = Join-Path $libPath "x64"
        $x86Path = Join-Path $libPath "x86"
        if (-not (Test-Path $x64Path)) { New-Item -ItemType Directory -Path $x64Path -Force | Out-Null }
        if (-not (Test-Path $x86Path)) { New-Item -ItemType Directory -Path $x86Path -Force | Out-Null }

        $sqliteDllDest = Join-Path $libPath "System.Data.SQLite.dll"
        
        # Determine current architecture for loading
        $currentArch = if ([IntPtr]::Size -eq 8) { "x64" } else { "x86" }
        $currentInteropPath = Join-Path $libPath "$currentArch\SQLite.Interop.dll"

        # Check if already installed (Managed DLL + Current Arch Native DLL)
        if ((Test-Path $sqliteDllDest) -and (Test-Path $currentInteropPath)) {
             Write-Verbose "System.Data.SQLite already installed for $currentArch."
             
             # Load it
             $env:PATH = "$(Join-Path $libPath $currentArch);" + $env:PATH
             [void][System.Reflection.Assembly]::LoadFrom($sqliteDllDest)
             try { Add-Type -Path $sqliteDllDest -ErrorAction SilentlyContinue } catch {}
             return $true
        }

        # Download Stub.System.Data.SQLite.Core.NetFramework which includes x64 binaries
        $nugetUrl = "https://www.nuget.org/api/v2/package/Stub.System.Data.SQLite.Core.NetFramework/1.0.118"
        $zipPath = Join-Path $libPath "sqlite.zip"
        $extractPath = Join-Path $libPath "sqlite"
        
        Write-Verbose "Downloading System.Data.SQLite from NuGet..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri $nugetUrl -OutFile $zipPath -UseBasicParsing -ErrorAction Stop
        
        # Extract the package
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        if (Test-Path $extractPath) { Remove-Item $extractPath -Recurse -Force }
        [System.IO.Compression.ZipFile]::ExtractToDirectory($zipPath, $extractPath)
        
        # Find source DLLs
        $sqliteDllSrc = Join-Path $extractPath "lib\net46\System.Data.SQLite.dll"
        $interopX64Src = Join-Path $extractPath "build\net46\x64\SQLite.Interop.dll"
        $interopX86Src = Join-Path $extractPath "build\net46\x86\SQLite.Interop.dll"
        
        if ((Test-Path $sqliteDllSrc) -and (Test-Path $interopX64Src)) {
            # Copy Managed DLL to root (try/catch in case locked)
            try { Copy-Item $sqliteDllSrc -Destination $libPath -Force -ErrorAction SilentlyContinue } catch {}
            
            # Copy Native DLLs to subfolders
            Copy-Item $interopX64Src -Destination $x64Path -Force
            if (Test-Path $interopX86Src) { Copy-Item $interopX86Src -Destination $x86Path -Force }
            
            # Set PATH for current architecture
            $env:PATH = "$(Join-Path $libPath $currentArch);" + $env:PATH
            
            # Load the DLL
            [void][System.Reflection.Assembly]::LoadFrom($sqliteDllDest)
            try { Add-Type -Path $sqliteDllDest -ErrorAction SilentlyContinue } catch {}
            
            Write-Verbose "✓ System.Data.SQLite installed and loaded successfully from NuGet"
            
            # Clean up temp files
            Remove-Item $zipPath -Force -ErrorAction SilentlyContinue
            Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
            
            return $true
        } else {
            Write-Verbose "✗ Could not find required DLLs in downloaded package"
            return $false
        }
    } catch {
        Write-Verbose "✗ Failed to download/install System.Data.SQLite: $_"
        return $false
    }
}

# =====================================================================================
# Load System.Data.SQLite (for Firefox HTML conversion)
# =====================================================================================
$script:SQLiteAvailable = $false

# Try to load from C# app's bin folder first (Only if running on Core/NET 5+, as these are likely NET 8 assemblies)
if ($PSVersionTable.PSEdition -eq 'Core') {
    $binPath = Join-Path $PSScriptRoot "bin\Release\net8.0\win-x64"
    $sqliteDllPath = Join-Path $binPath "System.Data.SQLite.dll"
    $interopDllPath = Join-Path $binPath "SQLite.Interop.dll"

    if ((Test-Path $sqliteDllPath) -and (Test-Path $interopDllPath)) {
        try {
            # Need to set the DLL directory for SQLite.Interop.dll dependency
            [System.Environment]::SetEnvironmentVariable('PATH', "$binPath;" + [System.Environment]::GetEnvironmentVariable('PATH'), 'Process')
            Add-Type -Path $sqliteDllPath
            $script:SQLiteAvailable = $true
            Write-Verbose "✓ Loaded System.Data.SQLite with dependencies from: $binPath"
        } catch {
            Write-Verbose "Failed to load from bin folder: $_"
        }
    }
}

# If not loaded from bin, try GAC
if (-not $script:SQLiteAvailable) {
    try {
        $loaded = [System.Reflection.Assembly]::LoadWithPartialName("System.Data.SQLite")
        if ($loaded) {
            $script:SQLiteAvailable = $true
            Write-Verbose "✓ Loaded System.Data.SQLite from GAC"
        }
    } catch {
        Write-Verbose "Not available in GAC: $_"
    }
}

# If still not loaded, try to download and install
if (-not $script:SQLiteAvailable) {
    Write-Verbose "System.Data.SQLite not found locally, attempting to download from NuGet..."
    $script:SQLiteAvailable = Install-SQLiteIfMissing
    if (-not $script:SQLiteAvailable) {
        Write-Verbose "⚠ System.Data.SQLite not available - Firefox HTML conversion will be disabled"
    }
}

# =====================================================================================
# GLOBALS & CONFIG
# =====================================================================================
$script:Config = $null
$script:OperationResults = @()
$script:StartTime = Get-Date

function Get-DefaultConfiguration {
    @{ DefaultPath = ""; PreferNetworkPath = $true; DefaultBrowsers = @('Chrome','Edge');
       AutoBackupBeforeImport = $true; VerifyFileIntegrity = $true;
       NetworkTimeoutSeconds = 3; MaxRetryAttempts = 3; RetryDelaySeconds = 1;
       DetailedLogging = $true; LogRetentionDays = 30;
       ShowProgressIndicator = $true; ConfirmOperations = $true;
       CreateBackupOnExport = $false; CompressBackups = $false }
}

function Get-Configuration {
    param([string]$ConfigFilePath)
    if (!$ConfigFilePath) { $ConfigFilePath = if ($ConfigPath) { $ConfigPath } else { Join-Path $env:USERPROFILE "BookmarkTool.config.json" } }
    if (Test-Path $ConfigFilePath) {
        try {
            Write-Verbose "Loading configuration from: $ConfigFilePath"
            $cfgObj = Get-Content $ConfigFilePath -Raw | ConvertFrom-Json
            $cfg = Get-DefaultConfiguration
            $cfgObj.PSObject.Properties | ForEach-Object { $cfg[$_.Name] = $_.Value }
            Write-Verbose "Configuration loaded"
            return $cfg
        } catch {
            Write-Warning "Failed to load configuration at ${ConfigFilePath}: $_"
            return Get-DefaultConfiguration
        }
    }
    else { Write-Verbose "No configuration file found at $ConfigFilePath, using defaults" }
    Get-DefaultConfiguration
}

function Save-Configuration {
    param([hashtable]$Config,[string]$ConfigFilePath)
    if (!$ConfigFilePath) { $ConfigFilePath = if ($ConfigPath) { $ConfigPath } else { Join-Path $env:USERPROFILE "BookmarkTool.config.json" } }
    try { $Config | ConvertTo-Json -Depth 5 | Set-Content $ConfigFilePath -Encoding UTF8; Write-Verbose "Config saved: $ConfigFilePath"; $true } catch { Write-Warning "Failed to save config: $_"; $false }
}

$script:Config = Get-Configuration

# =====================================================================================
# LOGGING
# =====================================================================================
function Get-LogFilePath {
    $path = if ($TargetPath) { $TargetPath } elseif ($Silent) { Get-HomeSharePath } else { $env:USERPROFILE }
    Join-Path $path "BookmarkTool.log"
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')][string]$Level = 'INFO'
    )
    $logFile = Get-LogFilePath
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fffK')
    $line = "$timestamp [$Level] $Message"
    $dir = Split-Path $logFile -Parent
    if (!(Test-Path -LiteralPath $dir -PathType Container)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    $line | Out-File -FilePath $logFile -Append -Encoding UTF8
    switch ($Level) { 'INFO' { Write-Information $Message -InformationAction Continue } 'WARN' { Write-Warning $Message } 'ERROR' { Write-Error $Message } 'DEBUG' { Write-Verbose $Message } }
}

function Invoke-LogRetention { try { $log = Get-LogFilePath; if (Test-Path $log) { $age = (Get-Date) - (Get-Item $log).CreationTime; if ($age.Days -gt $script:Config.LogRetentionDays) { Remove-Item $log -Force } } } catch { } }

# =====================================================================================
# PREREQS & UTILITIES
# =====================================================================================
function Test-Prerequisites {
    param([string]$TargetPath)
    $issues = @()
    if ($PSVersionTable.PSVersion.Major -lt 5) { $issues += 'PowerShell 5.1 or higher required' }
    if ($TargetPath) {
        try {
            if (!(Test-Path $TargetPath -IsValid)) { $issues += "Invalid target path format: $TargetPath" }
            if (Test-Path $TargetPath) {
                $testFile = Join-Path $TargetPath "_bmtool_permission_test.tmp"
                try { New-Item -Path $testFile -ItemType File -Force | Out-Null; Remove-Item -Path $testFile -Force | Out-Null }
                catch { $issues += "No write permission to target path: $TargetPath" }
            }
        } catch { $issues += "Cannot access target path: $TargetPath - $_" }
    }
    try { Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop; Add-Type -AssemblyName System.Drawing -ErrorAction Stop } catch { $issues += 'Required .NET assemblies not available' }
    if ($issues.Count) { Write-Error 'Prerequisites check failed:'; $issues | ForEach-Object { Write-Error "  - $_" }; return $false }
    Write-Verbose 'All prerequisites satisfied'; $true
}

function Test-PathAccess {
    param([Parameter(Mandatory)][string]$Path,[switch]$RequireWrite)
    try {
        if (!(Test-Path -LiteralPath $Path)) { Write-Verbose "Path does not exist: $Path"; return $false }
        if ($RequireWrite) {
            $testFile = Join-Path $Path "_bmtool_write_test_$(Get-Random).tmp"
            try { New-Item -Path $testFile -ItemType File -Force | Out-Null; Remove-Item -Path $testFile -Force | Out-Null; Write-Verbose "Write access confirmed: $Path" } catch { Write-Verbose "No write access: $Path"; return $false }
        }
        return $true
    } catch { Write-Verbose "Failed to access ${Path}: $_"; $false }
}

function Invoke-WithRetry {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock,[int]$MaxAttempts = $script:Config.MaxRetryAttempts,[int]$InitialDelay = $script:Config.RetryDelaySeconds,[double]$BackoffMultiplier = 2.0)
    $attempt = 1; $delay = $InitialDelay; $lastError = $null
    while ($attempt -le $MaxAttempts) {
        try { Write-Log "Attempt $attempt of $MaxAttempts" 'DEBUG'; return & $ScriptBlock }
        catch { $lastError = $_; Write-Log "Attempt $attempt failed: $_" 'DEBUG'; if ($attempt -lt $MaxAttempts) { Write-Log "Waiting $delay sec before retry" 'DEBUG'; Start-Sleep $delay; $delay = [math]::Min($delay*$BackoffMultiplier,30) }; $attempt++ }
    }
    throw $lastError
}

function Select-FileDialog { param([string]$Filter = 'All files (*.*)|*.*'); $ofd = New-Object System.Windows.Forms.OpenFileDialog; $ofd.Filter = $Filter; $res = $ofd.ShowDialog(); if ($res -eq [System.Windows.Forms.DialogResult]::OK) { return $ofd.FileName } $null }

function Get-HomeSharePath {
    $desktopPath = [IO.Path]::Combine($env:USERPROFILE,'Desktop')
    $networkPath = if ($env:HOMESHARE) { $env:HOMESHARE } else { "\\server\home\$env:USERNAME" }
    if (!($networkPath -like "\\*")) { Write-Verbose 'Network path not UNC; using Desktop'; return $desktopPath }
    try {
        $result = Invoke-WithRetry -ScriptBlock {
            $job = Start-Job -ScriptBlock {
                param($Path)
                try { 
                    if (Test-Path -Path $Path -PathType Container) { 
                        $temp = Join-Path $Path "_bmtool_temp_$(Get-Random).txt"
                        New-Item -Path $temp -ItemType File -Force | Out-Null
                        Remove-Item -Path $temp -Force | Out-Null
                        return $true 
                    } 
                } catch { }
                return $false
            } -ArgumentList $networkPath
            $isAccessible = $false; $timeout = $script:Config.NetworkTimeoutSeconds
            if (Wait-Job -Job $job -Timeout $timeout) { $isAccessible = Receive-Job -Job $job } else { Write-Log "Network path check timed out after $timeout s" 'DEBUG'; Stop-Job -Job $job -ErrorAction SilentlyContinue }
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            if (!$isAccessible) { throw "Network path not accessible: $networkPath" }
            $networkPath
        }
        Write-Verbose "Using network path: $result"; return $result
    } catch { Write-Verbose "Network path failed after retries: $_"; Write-Verbose "Using Desktop fallback: $desktopPath"; return $desktopPath }
}

# =====================================================================================
# BROWSER DETECTION & PROFILES
# =====================================================================================
function Test-BrowserRunning { param([string]$Browser)
    $processMap = @{ 'chrome'=@('chrome','GoogleChromeHelper','Google Chrome Helper','Google Chrome Helper (Renderer)','Google Chrome Helper (GPU)','Google Chrome Helper (Plugin)','crashpad_handler'); 'msedge'=@('msedge','MicrosoftEdge','MicrosoftEdgeWebView2','msedgewebview2','MicrosoftEdgeCP','MicrosoftEdgeSH','identity_helper'); 'firefox'=@('firefox','plugin-container','firefox.exe','crashreporter','updater','maintenanceservice') }
    $key = $Browser.ToLower(); if (-not $processMap.ContainsKey($key)) { Write-Warning "Unknown browser: $Browser"; return $false }
    $found = @(); foreach ($n in $processMap[$key]) { $p = Get-Process -Name $n -ErrorAction SilentlyContinue; if ($p) { $found += $p; Write-Verbose "Found running: $n (PID: $($p.Id -join ', '))" } }
    if ($found.Count -gt 0) { Write-Verbose "$Browser running with $($found.Count) related processes"; return $true } else { Write-Verbose "$Browser not running"; return $false }
}

function Close-Browser {
    param([Parameter(Mandatory)][string]$Browser)
    Write-Log "Attempting to close $Browser..."
    $processMap = @{ 'chrome'=@('chrome','GoogleChromeHelper','crashpad_handler'); 'edge'=@('msedge','MicrosoftEdge','MicrosoftEdgeWebView2','identity_helper'); 'firefox'=@('firefox','plugin-container','crashreporter') }
    $key = $Browser.ToLower(); if (-not $processMap.ContainsKey($key)) { Write-Warning "Unknown browser: $Browser"; return $false }
    
    $anyClosed = $false
    foreach ($processName in $processMap[$key]) {
        $processes = Get-Process -Name $processName -ErrorAction SilentlyContinue
        foreach ($proc in $processes) {
            try {
                Write-Verbose "Closing process: $processName (PID: $($proc.Id))"
                # Try graceful close first
                $proc.CloseMainWindow() | Out-Null
                $proc.WaitForExit(3000) | Out-Null
                
                # Force kill if still running
                if (-not $proc.HasExited) {
                    Write-Verbose "Force killing: $processName (PID: $($proc.Id))"
                    Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
                }
                $anyClosed = $true
            } catch {
                Write-Verbose "Failed to close $processName : $_"
            }
        }
    }
    
    if ($anyClosed) {
        Write-Log "$Browser closed successfully"
        return $true
    } else {
        Write-Log "Failed to close $Browser or it wasn't running" 'WARN'
        return $false
    }
}

function Get-AllBrowserProfiles {
    param([Parameter(Mandatory)][string]$Browser,[Parameter(Mandatory)][string]$FileName)
    $profiles = @()
    switch ($Browser.ToLower()) {
        'chrome' {
            $base = "$env:LOCALAPPDATA\Google\Chrome\User Data"
            if (Test-Path $base) {
                $dirs = Get-ChildItem -Path $base -Directory | Where-Object { ($_.Name -eq 'Default' -or $_.Name -match '^Profile \d+$') -and (Test-Path (Join-Path $_.FullName $FileName)) }
                foreach ($d in $dirs) { $profiles += @{ Name = if ($d.Name -eq 'Default') { 'Default Profile' } else { $d.Name }; Path = $d.FullName; LastUsed = $d.LastWriteTime } }
            }
        }
        'edge' {
            $base = "$env:LOCALAPPDATA\Microsoft\Edge\User Data"
            if (Test-Path $base) {
                $dirs = Get-ChildItem -Path $base -Directory | Where-Object { ($_.Name -eq 'Default' -or $_.Name -match '^Profile \d+$') -and (Test-Path (Join-Path $_.FullName $FileName)) }
                foreach ($d in $dirs) { $profiles += @{ Name = if ($d.Name -eq 'Default') { 'Default Profile' } else { $d.Name }; Path = $d.FullName; LastUsed = $d.LastWriteTime } }
            }
        }
        'firefox' {
            $base = "$env:APPDATA\Mozilla\Firefox\Profiles"
            if (Test-Path $base) {
                $dirs = Get-ChildItem -Path $base -Directory | Where-Object { Test-Path (Join-Path $_.FullName $FileName) }
                foreach ($d in $dirs) { $profiles += @{ Name = $d.Name; Path = $d.FullName; LastUsed = $d.LastWriteTime } }
            }
        }
    }
    $profiles | Sort-Object LastUsed -Descending
}

function Get-LatestProfilePath { param([string]$BasePath,[string]$FileName)
    if (!(Test-Path $BasePath)) { Write-Verbose "Base path not found: $BasePath"; return $null }
    $profiles = Get-ChildItem -Path $BasePath -Directory -ErrorAction SilentlyContinue | Where-Object { Test-Path (Join-Path $_.FullName $FileName) } | Sort-Object LastWriteTime -Descending
    if ($profiles) { Write-Verbose "Found profile: $($profiles[0].FullName)"; return $profiles[0].FullName } else { Write-Verbose "No profiles with $FileName in $BasePath"; return $null }
}

function Get-ChromeProfile { Get-LatestProfilePath "$env:LOCALAPPDATA\Google\Chrome\User Data" 'Bookmarks' }
function Get-EdgeProfile   { Get-LatestProfilePath "$env:LOCALAPPDATA\Microsoft\Edge\User Data"  'Bookmarks' }

function Get-FirefoxProfile {
    $ini = "$env:APPDATA\Mozilla\Firefox\profiles.ini"
    if (!(Test-Path $ini)) { Write-Verbose "Firefox profiles.ini not found: $ini"; return $null }
    $lines = Get-Content $ini | Where-Object { $_ -and ($_ -notmatch '^\s*#') }
    $pathLine  = $lines | Where-Object { $_ -match '^Path=' } | Select-Object -First 1
    $isRelLine = $lines | Where-Object { $_ -match '^IsRelative=' } | Select-Object -First 1
    $isRel = $false; if ($isRelLine) { $isRel = ($isRelLine -split '=',2)[1] -eq '1' }
    $rawPath = if ($pathLine) { ($pathLine -split '=',2)[1] } else { $null }
    if ($rawPath) {
        $profileFull = if ($isRel) { Join-Path "$env:APPDATA\Mozilla\Firefox" $rawPath } else { $rawPath }
        if (Test-Path (Join-Path $profileFull 'places.sqlite')) { Write-Verbose "Found Firefox profile: $profileFull"; return $profileFull }
    }
    Write-Verbose 'Primary profile not valid; searching for any profile with places.sqlite'
    Get-LatestProfilePath "$env:APPDATA\Mozilla\Firefox\Profiles" 'places.sqlite'
}

# =====================================================================================
# BACKUP & INTEGRITY
# =====================================================================================
function Backup-ExistingBookmarks {
    param([Parameter(Mandatory)][string]$BrowserProfile,[Parameter(Mandatory)][string]$BookmarkFile,[Parameter(Mandatory)][string]$BrowserName)
    try {
        $sourceFile = Join-Path $BrowserProfile $BookmarkFile
        if (!(Test-Path $sourceFile)) { Write-Verbose "No existing $BrowserName bookmark file to backup"; return $null }
        $backupDir = Join-Path $BrowserProfile 'BookmarkTool_Backups'
        if (!(Test-Path $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
        $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $backupPath = Join-Path $backupDir ("$BookmarkFile.$timestamp.backup")
        Copy-Item $sourceFile $backupPath -Force
        Write-Log "Created backup: $backupPath"
        $backups = Get-ChildItem $backupDir -Filter "$BookmarkFile.*.backup" | Sort-Object LastWriteTime -Descending
        if ($backups.Count -gt 10) { $backups | Select-Object -Skip 10 | ForEach-Object { Remove-Item $_.FullName -Force; Write-Verbose "Removed old backup: $($_.Name)" } }
        $backupPath
    } catch { Write-Warning "Failed to create backup for ${BrowserName}: $_"; $null }
}

function Test-BookmarkFileIntegrity {
    param([Parameter(Mandatory)][string]$FilePath,[Parameter(Mandatory)][string]$BrowserType)
    if (!(Test-Path $FilePath)) { Write-Verbose "File does not exist: $FilePath"; return $false }
    try {
        switch ($BrowserType.ToLower()) {
            'chrome' {
                $content = Get-Content $FilePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                $isValid = ($null -ne $content.roots) -and ($null -ne $content.version) -and ($null -ne $content.roots.bookmark_bar)
                if ($isValid) { Write-Verbose "Chrome bookmark file valid: $FilePath" } else { Write-Verbose 'Chrome file missing required structure' }
                return $isValid
            }
            'edge' {
                $content = Get-Content $FilePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                $isValid = ($null -ne $content.roots) -and ($null -ne $content.version)
                if ($isValid) { Write-Verbose "Edge bookmark file valid: $FilePath" } else { Write-Verbose 'Edge file missing required structure' }
                return $isValid
            }
            'firefox' {
                $fi = Get-Item $FilePath
                $isValid = ($fi.Length -gt 0) -and ($fi.Extension -eq '.sqlite')
                if ($isValid) {
                    $header = [System.IO.File]::ReadAllBytes($FilePath) | Select-Object -First 16
                    $sig = [System.Text.Encoding]::ASCII.GetString($header[0..15])
                    $isValid = $sig.StartsWith('SQLite format 3')
                }
                if ($isValid) { Write-Verbose "Firefox sqlite valid: $FilePath" } else { Write-Verbose 'Invalid Firefox sqlite format' }
                return $isValid
            }
            default { Write-Warning "Unknown browser type for integrity check: $BrowserType"; return $false }
        }
    } catch { Write-Verbose "Integrity check failed for ${FilePath}: $_"; return $false }
}

function New-OperationSummary {
    param([array]$Operations,[datetime]$StartTime,[string]$OperationType)
    $endTime = Get-Date; $duration = $endTime - $StartTime
    $successful = @($Operations | Where-Object { $_.Success })
    $failed     = @($Operations | Where-Object { -not $_.Success })
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
    foreach ($op in $Operations) { $status = if ($op.Success) { '[SUCCESS]' } else { '[FAILED]' }; $summary += "`n$status - $($op.Browser): $($op.Message)" }
    $summary += "`n═══════════════════════════════════════════════════════════════"
    $summary
}

# =====================================================================================
# HTML CONVERSION FUNCTIONS
# =====================================================================================
function ConvertTo-ChromeHtml {
    param([Parameter(Mandatory)][string]$JsonPath,[Parameter(Mandatory)][string]$OutputPath)
    try {
        $json = Get-Content $JsonPath -Raw | ConvertFrom-Json
        $html = @"
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<!-- This is an automatically generated file. It will be read and overwritten. DO NOT EDIT! -->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>
<DL><p>
"@
        
        function Add-BookmarkFolder {
            param($folder, $indent = 1)
            $spaces = "    " * $indent
            $result = "$spaces<DT><H3>$($folder.name)</H3>`n$spaces<DL><p>`n"
            
            if ($folder.children) {
                foreach ($child in $folder.children) {
                    if ($child.type -eq 'folder') {
                        $result += Add-BookmarkFolder $child ($indent + 1)
                    } elseif ($child.type -eq 'url') {
                        $addDate = if ($child.date_added) { " ADD_DATE=`"$($child.date_added)`"" } else { "" }
                        $result += "$spaces    <DT><A HREF=`"$($child.url)`"$addDate>$($child.name)</A>`n"
                    }
                }
            }
            $result += "$spaces</DL><p>`n"
            return $result
        }
        
        if ($json.roots.bookmark_bar) {
            $html += Add-BookmarkFolder $json.roots.bookmark_bar
        }
        if ($json.roots.other) {
            $html += Add-BookmarkFolder $json.roots.other
        }
        
        $html += "</DL><p>`n"
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Log "Converted Chrome bookmarks to HTML: $OutputPath"
        return $true
    } catch {
        Write-Log "Failed to convert Chrome bookmarks to HTML: $_" 'ERROR'
        return $false
    }
}

function ConvertTo-FirefoxHtml {
    param([Parameter(Mandatory)][string]$SqlitePath,[Parameter(Mandatory)][string]$OutputPath)
    
    if (-not $script:SQLiteAvailable) {
        Write-Log "System.Data.SQLite not available - Firefox HTML conversion skipped" 'INFO'
        Write-Log "Firefox bookmarks exported as SQLite database successfully" 'INFO'
        return $false
    }
    
    $connection = $null
    $tempDbPath = $null

    try {
        # Copy to temp file to avoid UNC path issues with SQLite
        $tempDbPath = [System.IO.Path]::GetTempFileName()
        Copy-Item -LiteralPath $SqlitePath -Destination $tempDbPath -Force

        # Create connection using direct type instantiation after assembly is loaded
        $connectionString = "Data Source=$tempDbPath;Version=3;Read Only=True;"
        
        # Try standard New-Object first
        try {
            $connection = New-Object -TypeName System.Data.SQLite.SQLiteConnection -ArgumentList $connectionString -ErrorAction Stop
        } catch {
            # Fallback: Use Reflection to instantiate if the type isn't visible to PowerShell
            Write-Verbose "Standard instantiation failed, trying reflection..."
            $assembly = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.Data.SQLite' } | Select-Object -First 1
            if ($null -eq $assembly) { throw "System.Data.SQLite assembly is not loaded." }
            
            $type = $assembly.GetType('System.Data.SQLite.SQLiteConnection')
            if ($null -eq $type) { throw "System.Data.SQLite.SQLiteConnection type not found in assembly." }
            
            $connection = [Activator]::CreateInstance($type, @($connectionString))
        }

        $connection.Open()
        
        $html = @"
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<!-- This is an automatically generated file. It will be read and overwritten. DO NOT EDIT! -->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>
<DL><p>
"@
        
        $query = "SELECT mb.id, mb.parent, mb.title, mp.url, mb.dateAdded FROM moz_bookmarks mb LEFT JOIN moz_places mp ON mb.fk = mp.id WHERE mb.type = 1 ORDER BY mb.parent, mb.position"
        $command = $connection.CreateCommand()
        $command.CommandText = $query
        $reader = $command.ExecuteReader()
        
        while ($reader.Read()) {
            $title = $reader["title"]
            $url = $reader["url"]
            $dateAdded = $reader["dateAdded"]
            if ($url) {
                $addDate = if ($dateAdded) { " ADD_DATE=`"$dateAdded`"" } else { "" }
                $html += "    <DT><A HREF=`"$url`"$addDate>$title</A>`n"
            }
        }
        
        $reader.Close()
        $connection.Close()
        
        $html += "</DL><p>`n"
        $html | Out-File -FilePath $OutputPath -Encoding UTF8
        Write-Log "Converted Firefox bookmarks to HTML: $OutputPath"
        return $true
    } catch {
        Write-Log "Failed to convert Firefox bookmarks to HTML: $_" 'WARN'
        if ($connection -and $connection.State -eq 'Open') { $connection.Close() }
        return $false
    } finally {
        if ($tempDbPath -and (Test-Path $tempDbPath)) {
            Remove-Item $tempDbPath -Force -ErrorAction SilentlyContinue
        }
    }
}

# =====================================================================================
# ZIP ARCHIVE FUNCTIONS
# =====================================================================================
function New-ZipArchive {
    param([Parameter(Mandatory)][string]$SourcePath,[Parameter(Mandatory)][string]$ZipPath)
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $files = Get-ChildItem -Path $SourcePath -File | Where-Object { $_.Extension -in '.json','.sqlite','.html','.htm' }
        
        if ($files.Count -eq 0) {
            Write-Log "No bookmark files found to zip in: $SourcePath" 'WARN'
            return $false
        }
        
        if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
        
        $zip = [System.IO.Compression.ZipFile]::Open($ZipPath, 'Create')
        foreach ($file in $files) {
            $entryName = $file.Name
            [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $file.FullName, $entryName) | Out-Null
            Write-Log "Added to ZIP: $entryName"
        }
        $zip.Dispose()
        
        Write-Log "Created ZIP archive: $ZipPath ($(Get-Item $ZipPath | Select-Object -ExpandProperty Length) bytes)"
        return $true
    } catch {
        Write-Log "Failed to create ZIP archive: $_" 'ERROR'
        return $false
    }
}

function Expand-ZipArchive {
    param([Parameter(Mandatory)][string]$ZipPath,[Parameter(Mandatory)][string]$DestinationPath)
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        
        if (-not (Test-Path $ZipPath)) {
            Write-Log "ZIP file not found: $ZipPath" 'ERROR'
            return $false
        }
        
        if (-not (Test-Path $DestinationPath)) {
            New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
        }
        
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $DestinationPath)
        Write-Log "Extracted ZIP archive to: $DestinationPath"
        return $true
    } catch {
        Write-Log "Failed to extract ZIP archive: $_" 'ERROR'
        return $false
    }
}

# =====================================================================================
# EXPORT / IMPORT
# =====================================================================================
function Export-Bookmarks {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Chrome,
        [switch]$Edge,
        [switch]$Firefox,
        [switch]$ExportHtmlOnly
    )

    if (!(Test-Path -LiteralPath $Path -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($Path,'Create backup directory')) { New-Item -ItemType Directory -Path $Path -Force | Out-Null; Write-Log "Created backup directory: $Path" }
    }

    $timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'

    function _Copy { param([string]$Src,[string]$Dst,[string]$Label)
        if ($PSCmdlet.ShouldProcess($Dst, "Export $Label bookmarks from $Src")) {
            Copy-Item -LiteralPath $Src -Destination $Dst -Force
            Write-Log "SUCCESS: Exported $Label from $Src to $Dst"
            $script:OperationResults += [pscustomobject]@{ Browser=$Label; Success=$true; Message="Exported to $Dst" }
        }
    }

    if ($Chrome) {
        Write-Log 'Starting Chrome export...'
        $profiles = if ($AllProfiles) { @(Get-AllBrowserProfiles -Browser 'Chrome' -FileName 'Bookmarks') } else { $p = Get-ChromeProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($profiles -and $profiles.Count -gt 0) { 
            foreach ($pr in $profiles) { 
                $src = Join-Path $pr.Path 'Bookmarks'
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\\s','_')" } else { '' }
                
                if ($ExportHtmlOnly) {
                    # Export only HTML
                    $dstHtml = Join-Path $Path ("Chrome$suffix`_Bookmarks_$timestamp.html")
                    if (ConvertTo-ChromeHtml -JsonPath $src -OutputPath $dstHtml) {
                        $script:OperationResults += [pscustomobject]@{ Browser='Chrome'; Success=$true; Message="Exported HTML to $dstHtml" }
                    } else {
                        $script:OperationResults += [pscustomobject]@{ Browser='Chrome'; Success=$false; Message="HTML conversion failed" }
                    }
                } else {
                    # Export both JSON and HTML (default)
                    $dstJson = Join-Path $Path ("Chrome$suffix`_BookmarkData_$timestamp.json")
                    $dstHtml = Join-Path $Path ("Chrome$suffix`_Bookmarks_$timestamp.html")
                    _Copy -Src $src -Dst $dstJson -Label 'Chrome'
                    ConvertTo-ChromeHtml -JsonPath $dstJson -OutputPath $dstHtml | Out-Null
                }
            }
        }
        else { Write-Log 'Chrome profile(s) not found - skipping Chrome export' 'WARN'; $script:OperationResults += [pscustomobject]@{ Browser='Chrome'; Success=$false; Message='Profile not found' } }
    }

    if ($Edge) {
        Write-Log 'Starting Edge export...'
        $profiles = if ($AllProfiles) { @(Get-AllBrowserProfiles -Browser 'Edge' -FileName 'Bookmarks') } else { $p = Get-EdgeProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($profiles -and $profiles.Count -gt 0) { 
            foreach ($pr in $profiles) { 
                $src = Join-Path $pr.Path 'Bookmarks'
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\\s','_')" } else { '' }
                
                if ($ExportHtmlOnly) {
                    # Export only HTML
                    $dstHtml = Join-Path $Path ("Edge$suffix`_Bookmarks_$timestamp.html")
                    if (ConvertTo-ChromeHtml -JsonPath $src -OutputPath $dstHtml) {
                        $script:OperationResults += [pscustomobject]@{ Browser='Edge'; Success=$true; Message="Exported HTML to $dstHtml" }
                    } else {
                        $script:OperationResults += [pscustomobject]@{ Browser='Edge'; Success=$false; Message="HTML conversion failed" }
                    }
                } else {
                    # Export both JSON and HTML (default)
                    $dstJson = Join-Path $Path ("Edge$suffix`_BookmarkData_$timestamp.json")
                    $dstHtml = Join-Path $Path ("Edge$suffix`_Bookmarks_$timestamp.html")
                    _Copy -Src $src -Dst $dstJson -Label 'Edge'
                    ConvertTo-ChromeHtml -JsonPath $dstJson -OutputPath $dstHtml | Out-Null
                }
            }
        }
        else { Write-Log 'Edge profile(s) not found - skipping Edge export' 'WARN'; $script:OperationResults += [pscustomobject]@{ Browser='Edge'; Success=$false; Message='Profile not found' } }
    }

    if ($Firefox) {
        Write-Log 'Starting Firefox export...'
        $profiles = if ($AllProfiles) { @(Get-AllBrowserProfiles -Browser 'Firefox' -FileName 'places.sqlite') } else { $p = Get-FirefoxProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($profiles -and $profiles.Count -gt 0) { 
            foreach ($pr in $profiles) { 
                $src = Join-Path $pr.Path 'places.sqlite'
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\\s','_')" } else { '' }
                
                if ($ExportHtmlOnly) {
                    # Export only HTML
                    $dstHtml = Join-Path $Path ("Firefox$suffix`_Bookmarks_$timestamp.html")
                    if (ConvertTo-FirefoxHtml -SqlitePath $src -OutputPath $dstHtml) {
                        $script:OperationResults += [pscustomobject]@{ Browser='Firefox'; Success=$true; Message="Exported HTML to $dstHtml" }
                    } else {
                        $script:OperationResults += [pscustomobject]@{ Browser='Firefox'; Success=$false; Message="HTML conversion failed" }
                    }
                } else {
                    # Export both SQLite and HTML (default)
                    $dstSqlite = Join-Path $Path ("Firefox$suffix`_BookmarkData_$timestamp.sqlite")
                    $dstHtml = Join-Path $Path ("Firefox$suffix`_Bookmarks_$timestamp.html")
                    _Copy -Src $src -Dst $dstSqlite -Label 'Firefox'
                    ConvertTo-FirefoxHtml -SqlitePath $dstSqlite -OutputPath $dstHtml | Out-Null
                }
            }
        }
        else { Write-Log 'Firefox profile(s) not found - skipping Firefox export' 'WARN'; $script:OperationResults += [pscustomobject]@{ Browser='Firefox'; Success=$false; Message='Profile not found' } }
    }

    Write-Log 'Export operation completed'
    
    # Create ZIP archive if requested
    if ($CreateZip) {
        $zipPath = Join-Path $Path "BookmarkBackup_$timestamp.zip"
        Write-Log "Creating ZIP archive: $zipPath"
        if (New-ZipArchive -SourcePath $Path -ZipPath $zipPath) {
            $script:OperationResults += [pscustomobject]@{ Browser='ZIP'; Success=$true; Message="Created archive: $zipPath" }
        } else {
            $script:OperationResults += [pscustomobject]@{ Browser='ZIP'; Success=$false; Message="ZIP creation failed" }
        }
    }
}

function Import-Bookmarks {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Path,[switch]$Chrome,[switch]$Edge,[switch]$Firefox,[switch]$CloseBrowserIfRunning)

    # Check if browsers are running and offer to close them
    $browsersToClose = @()
    if ($Chrome -and (Test-BrowserRunning 'chrome')) { $browsersToClose += 'Chrome' }
    if ($Edge -and (Test-BrowserRunning 'msedge')) { $browsersToClose += 'Edge' }
    if ($Firefox -and (Test-BrowserRunning 'firefox')) { $browsersToClose += 'Firefox' }
    
    if ($browsersToClose.Count -gt 0) {
        $browserList = $browsersToClose -join ', '
        Write-Log "WARNING: The following browsers are running: $browserList" 'WARN'
        
        if ($CloseBrowserIfRunning -or $Force) {
            Write-Log "Attempting to close running browsers..."
            foreach ($browser in $browsersToClose) {
                $browserKey = switch ($browser) {
                    'Chrome' { 'chrome' }
                    'Edge' { 'edge' }
                    'Firefox' { 'firefox' }
                }
                Close-Browser -Browser $browserKey
            }
            Start-Sleep -Seconds 2
            Write-Log "Browsers closed. Proceeding with import..."
        } elseif (-not $Silent) {
            Write-Host "`n⚠️ WARNING: $browserList is currently running!" -ForegroundColor Yellow
            Write-Host "The browser(s) must be closed to import bookmarks safely.`n" -ForegroundColor Yellow
            Write-Host "Would you like to close $browserList now? (Y/N): " -NoNewline -ForegroundColor Cyan
            $response = Read-Host
            if ($response -match '^[Yy]') {
                foreach ($browser in $browsersToClose) {
                    $browserKey = switch ($browser) {
                        'Chrome' { 'chrome' }
                        'Edge' { 'edge' }
                        'Firefox' { 'firefox' }
                    }
                    Close-Browser -Browser $browserKey
                }
                Start-Sleep -Seconds 2
                Write-Log "Browsers closed. Proceeding with import..."
            } else {
                Write-Log "Import cancelled by user - browsers still running" 'WARN'
                return
            }
        } else {
            Write-Log "ABORT: Browsers running in silent mode without -CloseBrowserIfRunning flag" 'ERROR'
            return
        }
    }

    function _CopyIn { param([string]$Src,[string]$Dst,[string]$Label)
        if (!(Test-Path -LiteralPath $Src)) { Write-Log "$Label import source not found: $Src" 'WARN'; $script:OperationResults += [pscustomobject]@{ Browser=$Label; Success=$false; Message="Source missing: $Src" }; return }
        if ($script:Config.AutoBackupBeforeImport) { $null = Backup-ExistingBookmarks -BrowserProfile (Split-Path $Dst -Parent) -BookmarkFile (Split-Path $Dst -Leaf) -BrowserName $Label }
        if ($PSCmdlet.ShouldProcess($Dst, "Import $Label bookmarks from $Src")) {
            Copy-Item -LiteralPath $Src -Destination $Dst -Force
            Write-Log "SUCCESS: Imported $Label from $Src to $Dst"
            $script:OperationResults += [pscustomobject]@{ Browser=$Label; Success=$true; Message="Imported from $Src" }
        }
    }

    if ($Chrome) {
        Write-Log 'Starting Chrome import...'
        $targets = if ($AllProfiles) { @(Get-AllBrowserProfiles -Browser 'Chrome' -FileName 'Bookmarks') } else { $p = Get-ChromeProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($targets -and $targets.Count -gt 0) {
            foreach ($pr in $targets) {
                $dst = Join-Path $pr.Path 'Bookmarks'
                $candidate = if ($AllProfiles) { Join-Path $Path ("Chrome-Bookmarks-$($pr.Name -replace '\\s','_').json") } else { Join-Path $Path 'Chrome-Bookmarks.json' }
                $src = if (Test-Path $candidate) { $candidate } else { Join-Path $Path 'Chrome-Bookmarks.json' }
                if (!(Test-Path $src) -and -not $Silent) { Write-Log "Prompting for Chrome file" 'INFO'; $tmp = Select-FileDialog -Filter 'Chrome Bookmarks (*.json)|*.json'; if ($tmp) { $src = $tmp; Write-Log "User selected Chrome import file: $src" } }
                _CopyIn -Src $src -Dst $dst -Label 'Chrome'
            }
        } else { Write-Log 'Chrome profile(s) not found' 'WARN' }
    }

    if ($Edge) {
        Write-Log 'Starting Edge import...'
        $targets = if ($AllProfiles) { @(Get-AllBrowserProfiles -Browser 'Edge' -FileName 'Bookmarks') } else { $p = Get-EdgeProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($targets -and $targets.Count -gt 0) {
            foreach ($pr in $targets) {
                $dst = Join-Path $pr.Path 'Bookmarks'
                $candidate = if ($AllProfiles) { Join-Path $Path ("Edge-Bookmarks-$($pr.Name -replace '\\s','_').json") } else { Join-Path $Path 'Edge-Bookmarks.json' }
                $src = if (Test-Path $candidate) { $candidate } else { Join-Path $Path 'Edge-Bookmarks.json' }
                if (!(Test-Path $src) -and -not $Silent) { Write-Log "Prompting for Edge file" 'INFO'; $tmp = Select-FileDialog -Filter 'Edge Bookmarks (*.json)|*.json'; if ($tmp) { $src = $tmp; Write-Log "User selected Edge import file: $src" } }
                _CopyIn -Src $src -Dst $dst -Label 'Edge'
            }
        } else { Write-Log 'Edge profile(s) not found' 'WARN' }
    }

    if ($Firefox) {
        Write-Log 'Starting Firefox import...'
        $targets = if ($AllProfiles) { @(Get-AllBrowserProfiles -Browser 'Firefox' -FileName 'places.sqlite') } else { $p = Get-FirefoxProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($targets -and $targets.Count -gt 0) {
            foreach ($pr in $targets) {
                $dst = Join-Path $pr.Path 'places.sqlite'
                $candidate = if ($AllProfiles) { Join-Path $Path ("Firefox-places-$($pr.Name -replace '\\s','_').sqlite") } else { Join-Path $Path 'Firefox-places.sqlite' }
                $src = if (Test-Path $candidate) { $candidate } else { Join-Path $Path 'Firefox-places.sqlite' }
                if (!(Test-Path $src) -and -not $Silent) { Write-Log "Prompting for Firefox file" 'INFO'; $tmp = Select-FileDialog -Filter 'Firefox places.sqlite|places.sqlite'; if ($tmp) { $src = $tmp; Write-Log "User selected Firefox import file: $src" } }
                _CopyIn -Src $src -Dst $dst -Label 'Firefox'
            }
        } else { Write-Log 'Firefox profile(s) not found' 'WARN' }
    }

    Write-Log 'Import operation completed'
}

function Import-FromZip {
    param([Parameter(Mandatory)][string]$ZipPath,[string]$TargetBrowser)
    
    $tempExtractPath = Join-Path $env:TEMP "BookmarkImport_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    
    try {
        Write-Log "Extracting ZIP archive: $ZipPath"
        if (-not (Expand-ZipArchive -ZipPath $ZipPath -DestinationPath $tempExtractPath)) {
            return $false
        }
        
        $extractedFiles = Get-ChildItem -Path $tempExtractPath -File
        Write-Log "Extracted $($extractedFiles.Count) files from ZIP"
        
        # Determine which browser files are present
        $chromeFiles = $extractedFiles | Where-Object { $_.Name -match 'Chrome.*\.json' }
        $edgeFiles = $extractedFiles | Where-Object { $_.Name -match 'Edge.*\.json' }
        $firefoxFiles = $extractedFiles | Where-Object { $_.Name -match 'Firefox.*\.sqlite' }
        
        $importSuccess = $false
        
        # Import based on target browser or auto-detect
        if ($TargetBrowser -eq 'Chrome' -and $chromeFiles) {
            Write-Log "Importing Chrome bookmarks from ZIP"
            Import-Bookmarks -Path $tempExtractPath -Chrome
            $importSuccess = $true
        } elseif ($TargetBrowser -eq 'Edge' -and $edgeFiles) {
            Write-Log "Importing Edge bookmarks from ZIP"
            Import-Bookmarks -Path $tempExtractPath -Edge
            $importSuccess = $true
        } elseif ($TargetBrowser -eq 'Firefox' -and $firefoxFiles) {
            Write-Log "Importing Firefox bookmarks from ZIP"
            Import-Bookmarks -Path $tempExtractPath -Firefox
            $importSuccess = $true
        } else {
            # Auto-detect and import all available
            Write-Log "Auto-detecting browsers in ZIP archive"
            if ($chromeFiles) { Import-Bookmarks -Path $tempExtractPath -Chrome; $importSuccess = $true }
            if ($edgeFiles) { Import-Bookmarks -Path $tempExtractPath -Edge; $importSuccess = $true }
            if ($firefoxFiles) { Import-Bookmarks -Path $tempExtractPath -Firefox; $importSuccess = $true }
        }
        
        return $importSuccess
    } catch {
        Write-Log "Failed to import from ZIP: $_" 'ERROR'
        return $false
    } finally {
        # Clean up temp extraction folder
        if (Test-Path $tempExtractPath) {
            Remove-Item -Path $tempExtractPath -Recurse -Force -ErrorAction SilentlyContinue
            Write-Log "Cleaned up temporary extraction folder"
        }
    }
}

# =====================================================================================
# GUI
# =====================================================================================
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-ProgressDialog {
    param([string]$Title = 'Processing...',[string]$Message = 'Please wait...',[scriptblock]$Operation)
    $form = New-Object Windows.Forms.Form
    $form.Text = $Title; $form.Size = '400,150'; $form.StartPosition = 'CenterParent'; $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.TopMost = $true
    $lbl = New-Object Windows.Forms.Label; $lbl.Text = $Message; $lbl.Location = '20,20'; $lbl.Size = '360,30'; $lbl.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter; $form.Controls.Add($lbl)
    $bar = New-Object Windows.Forms.ProgressBar; $bar.Location = '20,60'; $bar.Size = '360,25'; $bar.Style = 'Marquee'; $bar.MarqueeAnimationSpeed = 50; $form.Controls.Add($bar)
    $operationResult = $null; $operationError = $null
    $timer = New-Object Windows.Forms.Timer; $timer.Interval = 100
    $job = Start-Job -ScriptBlock $Operation
    $timer.Add_Tick({ if ($job.State -eq 'Completed') { $timer.Stop(); try { $script:operationResult = Receive-Job -Job $job } catch { $script:operationError = $_ }; Remove-Job -Job $job -Force; $form.Close() } elseif ($job.State -in ('Failed','Stopped')) { $timer.Stop(); try { $script:operationError = Receive-Job -Job $job } catch { $script:operationError = $_ }; Remove-Job -Job $job -Force; $form.Close() } })
    $timer.Start(); $form.ShowDialog() | Out-Null; $timer.Stop()
    if ($operationError) { throw $operationError }
    $operationResult
}

function Show-GUI {
    Write-Log 'Launching GUI mode'
    $form = New-Object Windows.Forms.Form
    $form.Text = 'Bookmark Backup Tool v5.2 - Enhanced Edition'; $form.Size = '600,480'; $form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false

    $chkChrome  = New-Object Windows.Forms.CheckBox; $chkChrome.Text='Chrome';  $chkChrome.Location='30,30';  $chkChrome.AutoSize=$true; $form.Controls.Add($chkChrome)
    $chkEdge    = New-Object Windows.Forms.CheckBox; $chkEdge.Text='Edge';      $chkEdge.Location='30,60';  $chkEdge.AutoSize=$true; $form.Controls.Add($chkEdge)
    $chkFirefox = New-Object Windows.Forms.CheckBox; $chkFirefox.Text='Firefox';$chkFirefox.Location='30,90'; $chkFirefox.AutoSize=$true; $form.Controls.Add($chkFirefox)
    
    $chkAllProfiles = New-Object Windows.Forms.CheckBox; $chkAllProfiles.Text='Process all profiles (not just latest)';$chkAllProfiles.Location='30,120'; $chkAllProfiles.AutoSize=$true; $form.Controls.Add($chkAllProfiles)
    $chkHtmlOnly = New-Object Windows.Forms.CheckBox; $chkHtmlOnly.Text='HTML-only export (lightweight)';$chkHtmlOnly.Location='30,145'; $chkHtmlOnly.AutoSize=$true; $form.Controls.Add($chkHtmlOnly)
    $chkCreateZip = New-Object Windows.Forms.CheckBox; $chkCreateZip.Text='Create ZIP archive';$chkCreateZip.Location='30,170'; $chkCreateZip.AutoSize=$true; $form.Controls.Add($chkCreateZip)

    $lblPath = New-Object Windows.Forms.Label; $lblPath.Text='Save/Load Path:'; $lblPath.Location='30,205'; $lblPath.Size='420,20'; $form.Controls.Add($lblPath)
    $txtPath = New-Object Windows.Forms.TextBox; $txtPath.Text = Get-HomeSharePath; $txtPath.Location='30,230'; $txtPath.Size='360,25'; $form.Controls.Add($txtPath)
    $btnBrowse = New-Object Windows.Forms.Button; $btnBrowse.Text='Browse'; $btnBrowse.Location='400,229'; $btnBrowse.Size='60,25'; $btnBrowse.Add_Click({ $dialog = New-Object Windows.Forms.FolderBrowserDialog; if ($dialog.ShowDialog() -eq 'OK') { $txtPath.Text = $dialog.SelectedPath; Write-Log "User selected path: $($dialog.SelectedPath)" } }); $form.Controls.Add($btnBrowse)

    $btnExport = New-Object Windows.Forms.Button; $btnExport.Text='Export Bookmarks'; $btnExport.Size='180,40'; $btnExport.Location='40,300'
    $btnExport.Add_Click({
        Write-Log 'Export button clicked'
        $currentPath = if ($txtPath.Text) { $txtPath.Text } else { Get-HomeSharePath }
        $txtPath.Text = $currentPath
        $script:AllProfiles = $chkAllProfiles.Checked
        $script:CreateZip = $chkCreateZip.Checked
        Export-Bookmarks -Path $currentPath -Chrome:$chkChrome.Checked -Edge:$chkEdge.Checked -Firefox:$chkFirefox.Checked -ExportHtmlOnly:$chkHtmlOnly.Checked
        [Windows.Forms.MessageBox]::Show('Export completed. Check the log for details.','Export Complete',[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        Write-Log 'Export completed via GUI'
    })
    $form.Controls.Add($btnExport)

    $btnImport = New-Object Windows.Forms.Button; $btnImport.Text='Import Bookmarks'; $btnImport.Size='180,40'; $btnImport.Location='260,300'
    $btnImport.Add_Click({
        Write-Log 'Import button clicked'
        $currentPath = if ($txtPath.Text) { $txtPath.Text } else { Get-HomeSharePath }
        $txtPath.Text = $currentPath
        $script:AllProfiles = $chkAllProfiles.Checked
        Import-Bookmarks -Path $currentPath -Chrome:$chkChrome.Checked -Edge:$chkEdge.Checked -Firefox:$chkFirefox.Checked -CloseBrowserIfRunning
        [Windows.Forms.MessageBox]::Show('Import completed. Check the log for details.','Import Complete',[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        Write-Log 'Import completed via GUI'
    })
    $form.Controls.Add($btnImport)
    
    $lblInfo = New-Object Windows.Forms.Label
    $lblInfo.Text = 'Tip: Export creates dual-format backups (native + HTML) by default'
    $lblInfo.Location = '30,360'
    $lblInfo.Size = '520,40'
    $lblInfo.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($lblInfo)

    Write-Log 'GUI displayed successfully'; $form.ShowDialog() | Out-Null; Write-Log 'GUI closed'
}

# =====================================================================================
# SCHEDULED TASKS
# =====================================================================================
function New-BookmarkScheduledTask {
    param([string]$Frequency = 'Daily',[string]$Time = '18:00')
    try {
        if (!(Get-Command Register-ScheduledTask -ErrorAction SilentlyContinue)) { Write-Warning 'Task Scheduler cmdlets not available.'; return $false }
        $scriptArgs = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Silent -Action Export -Chrome -Edge -Firefox"
        $action   = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument $scriptArgs
        switch ($Frequency.ToLower()) {
            'daily'   { $trigger = New-ScheduledTaskTrigger -Daily -At $Time }
            'weekly'  { $trigger = New-ScheduledTaskTrigger -Weekly -At $Time -DaysOfWeek Sunday }
            'monthly' { $trigger = New-ScheduledTaskTrigger -Monthly -DaysOfMonth 1 -At $Time }
            default   { throw "Invalid frequency: $Frequency. Must be Daily, Weekly, or Monthly." }
        }
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

        $taskName = 'BookmarkBackupTool_AutoExport'
        $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existing) { Write-Log "Scheduled task exists: $taskName"; Unregister-ScheduledTask -TaskName $taskName -Confirm:$false; Write-Log "Removed existing scheduled task: $taskName" }
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description 'Automatic bookmark backup using BookmarkTool v5.0' | Out-Null
        Write-Log "Successfully created scheduled task: $taskName"; Write-Log "Frequency: $Frequency at $Time"; Write-Information "[SUCCESS] Scheduled task created. Next Run: $($trigger.StartBoundary)"; $true
    } catch { Write-Error "Failed to create scheduled task: $_"; Write-Log "ERROR: Failed to create scheduled task - $_" 'ERROR'; $false }
}

function Remove-BookmarkScheduledTask { try { $name='BookmarkBackupTool_AutoExport'; if (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $name -Confirm:$false; Write-Log "Removed scheduled task: $name"; $true } else { Write-Warning "Scheduled task not found: $name"; $false } } catch { Write-Error "Failed to remove scheduled task: $_"; Write-Log "ERROR: Failed to remove scheduled task - $_" 'ERROR'; $false } }

# =====================================================================================
# MAIN EXECUTION LOGIC
# =====================================================================================
if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch '^\s*\.\s') {
    Invoke-LogRetention
    Write-Log '=== Bookmark Backup Tool v5.0 Enhanced Edition Started ==='
    Write-Log "PowerShell Version: $($PSVersionTable.PSVersion)"
    $execMode = 'GUI'
if ($CreateScheduledTask) { $execMode = 'Scheduled Task Creation' }
elseif ($Silent) { $execMode = 'Silent' }
Write-Log ("Execution mode: {0}" -f $execMode)
    try {
        Write-Verbose 'Checking system prerequisites...'
        if (-not (Test-Prerequisites -TargetPath $TargetPath)) { throw 'Prerequisites check failed.' }

        # STA relaunch only for GUI mode
        if (-not $Silent -and [Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
            Write-Verbose 'Re-launching in STA mode for Windows Forms...'
            $argsJoined = ($MyInvocation.UnboundArguments + $PSBoundParameters.GetEnumerator() | ForEach-Object {
                if ($_.GetType().Name -eq 'DictionaryEntry') {
                    if ($_.Value -is [switch] -and $_.Value.IsPresent) { "-$( $_.Key )" }
                    elseif ($_.Value -is [string]) { "-$( $_.Key ) `"$( $_.Value )`"" }
                    else { "-$( $_.Key ) $( $_.Value )" }
                } else { $_ }
            }) -join ' '
            Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList "-STA -NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $argsJoined" | Out-Null
            $global:LASTEXITCODE = 0; exit 0
        }

        if ($CreateScheduledTask) {
            Write-Log 'Creating scheduled task for automatic backups'
            if (New-BookmarkScheduledTask -Frequency $ScheduleFrequency) { Write-Log 'Scheduled task creation completed successfully'; $global:LASTEXITCODE = 0; exit 0 } else { throw 'Scheduled task creation failed.' }
        }

        $finalPath = if ($TargetPath) {
            Write-Log "Using specified target path: $TargetPath"
            if (!(Test-Path -LiteralPath $TargetPath -PathType Container)) { if ($PSCmdlet.ShouldProcess($TargetPath,'Create target directory')) { New-Item -ItemType Directory -Path $TargetPath -Force | Out-Null } }
            $TargetPath
        } elseif ($Silent) { $auto = Get-HomeSharePath; Write-Log "Auto-detected path for silent mode: $auto"; $auto } else { Write-Log 'GUI mode - path via user selection'; $null }

        if ($Silent) {
            Write-Log 'Executing in silent mode'
            if (-not $Chrome -and -not $Edge -and -not $Firefox) { Write-Log 'No browsers selected. Using defaults from configuration.' 'WARN'; $Chrome = $script:Config.DefaultBrowsers -contains 'Chrome'; $Edge = $script:Config.DefaultBrowsers -contains 'Edge'; $Firefox = $script:Config.DefaultBrowsers -contains 'Firefox' }
            $script:OperationResults = @(); $operationStart = Get-Date
            switch ($Action) {
                'Export' {
                    Write-Log "Silent export: Chrome=$Chrome Edge=$Edge Firefox=$Firefox HtmlOnly=$HtmlOnly; Path=$finalPath"
                    if ($PSCmdlet.ShouldProcess($finalPath,'Export bookmarks')) { Export-Bookmarks -Path $finalPath -Chrome:$Chrome -Edge:$Edge -Firefox:$Firefox -ExportHtmlOnly:$HtmlOnly }
                    $summary = New-OperationSummary -Operations $script:OperationResults -StartTime $operationStart -OperationType 'Export'
                    Write-Log $summary; Write-Information $summary
                }
                'Import' {
                    Write-Log "Silent import: Chrome=$Chrome Edge=$Edge Firefox=$Firefox; Path=$finalPath"
                    if ($PSCmdlet.ShouldProcess($finalPath,'Import bookmarks')) { Import-Bookmarks -Path $finalPath -Chrome:$Chrome -Edge:$Edge -Firefox:$Firefox }
                    $summary = New-OperationSummary -Operations $script:OperationResults -StartTime $operationStart -OperationType 'Import'
                    Write-Log $summary; Write-Information $summary
                }
                default { throw "When using -Silent, -Action must be 'Export' or 'Import'." }
            }
            Write-Log '=== Silent mode execution completed successfully ==='
            $global:LASTEXITCODE = 0; exit 0
        } else {
            Write-Log 'Launching enhanced GUI mode'
            Show-GUI
            Write-Log '=== GUI mode execution completed successfully ==='
            $global:LASTEXITCODE = 0; exit 0
        }
    }
    catch {
        Write-Log "ERROR: $($_.Exception.Message)" 'ERROR'
        $global:LASTEXITCODE = 1
        exit 1
    }
}
