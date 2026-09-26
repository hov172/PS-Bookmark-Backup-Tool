# =====================================================================================
# Bookmark Backup Tool v5.4 - Enhanced Edition
# Author: Jesus M. Ayala
# Version: 5.4  (also set in $script:ToolVersion below - keep both in sync)
# Last Modified: September 26th, 2026
# Requires: PowerShell 5.1+, Windows 10/11, .NET Framework (for System.Data.SQLite)
# License: MIT
# Encoding: this file must stay UTF-8 *with BOM* - Windows PowerShell 5.1 misreads it otherwise
#
# NEW IN v5.4 (reliability release):
# - Firefox: backups include unsaved changes (places.sqlite-wal merged in), even while Firefox is open;
#   every Firefox backup is a single standalone .sqlite file; works on network shares
# - Firefox: uses the profile Firefox actually launches ([Install*] Default= in profiles.ini)
# - Import: finds this tool's own exports (newest *_BookmarkData_<timestamp> file), validates the file
#   before overwriting (bad/truncated/wrong-browser files are skipped), clears stale Firefox WAL/SHM files,
#   and with -AllProfiles never gives one profile another profile's bookmarks
# - HTML export: keeps folders, escapes special characters, dates in seconds, toolbar marked for re-import
# - Silent mode without -TargetPath no longer hangs (log path / share probe recursion); share probe is fast
#   and reliable on PowerShell 5.1; HOMESHARE unset -> Desktop immediately; Desktop follows OneDrive redirection
# - Config DefaultPath and PreferNetworkPath are now honored
# - Pre-import backup rotation keeps the newest 10 by name (file dates are copied from the source)
# - A logged error no longer ends the run; ZIP holds only this run's files; summary printed once
# - -AllProfiles file names use underscores for spaces (older names still import)
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
# Load System.Data.SQLite (Firefox: WAL merge into backups, import validation, HTML conversion)
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
        Write-Verbose "⚠ System.Data.SQLite not available - Firefox HTML conversion, WAL merge and deep import validation will be disabled"
    }
}

# =====================================================================================
# GLOBALS & CONFIG
# =====================================================================================
$script:ToolVersion = '5.4'
# $true in the module built by Build-Module.ps1; $false when run as a script
$script:IsModule = $false
$script:ScriptPath = $PSCommandPath
$script:Config = $null
$script:HomeSharePathCache = $null
$script:ResolvingHomeShare = $false
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
    # $null while the home share is still being probed: those few lines go to the console only
    if (-not $TargetPath -and $Silent -and $script:ResolvingHomeShare) { return $null }
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
    if ($logFile) {
        $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fffK')
        $line = "$timestamp [$Level] $Message"
        $dir = Split-Path $logFile -Parent
        # -WhatIf:$false: the log is always written, even during a -WhatIf preview
        if (!(Test-Path -LiteralPath $dir -PathType Container)) { New-Item -ItemType Directory -Path $dir -Force -WhatIf:$false | Out-Null }
        $line | Out-File -FilePath $logFile -Append -Encoding UTF8 -WhatIf:$false
    }
    switch ($Level) { 'INFO' { Write-Information $Message -InformationAction Continue } 'WARN' { Write-Warning $Message } 'ERROR' { Write-Error $Message -ErrorAction Continue } 'DEBUG' { Write-Verbose $Message } }
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
                try { New-Item -Path $testFile -ItemType File -Force -WhatIf:$false | Out-Null; Remove-Item -Path $testFile -Force -WhatIf:$false | Out-Null }
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
            try { New-Item -Path $testFile -ItemType File -Force -WhatIf:$false | Out-Null; Remove-Item -Path $testFile -Force -WhatIf:$false | Out-Null; Write-Verbose "Write access confirmed: $Path" } catch { Write-Verbose "No write access: $Path"; return $false }
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
    # Probed once per run and cached: every Write-Log in silent mode asks for this path, and the probe itself
    # logs (via Invoke-WithRetry). Without the cache and the in-progress flag, that recursed forever.
    if ($script:HomeSharePathCache) { return $script:HomeSharePathCache }
    $script:ResolvingHomeShare = $true
    try { $script:HomeSharePathCache = Resolve-HomeSharePath } finally { $script:ResolvingHomeShare = $false }
    $script:HomeSharePathCache
}

function Resolve-HomeSharePath {
    # Order: config DefaultPath (if usable) -> HOMESHARE network share (unless PreferNetworkPath is false) -> Desktop.
    # Windows' real Desktop location (follows OneDrive/folder redirection); %USERPROFILE%\Desktop only if that fails
    $desktopPath = [Environment]::GetFolderPath('Desktop')
    if (-not $desktopPath) { $desktopPath = [IO.Path]::Combine($env:USERPROFILE,'Desktop') }

    $configured = if ($script:Config) { [string]$script:Config.DefaultPath } else { '' }
    if ($configured) {
        $configured = [Environment]::ExpandEnvironmentVariables($configured)
        try {
            if (!(Test-Path -LiteralPath $configured -PathType Container)) { New-Item -ItemType Directory -Path $configured -Force | Out-Null }
            Write-Verbose "Using configured DefaultPath: $configured"; return $configured
        } catch { Write-Verbose "Configured DefaultPath not usable ($configured): $_ - falling back to auto-detection" }
    }
    if ($script:Config -and $script:Config.PreferNetworkPath -eq $false) { Write-Verbose 'PreferNetworkPath is false; using Desktop'; return $desktopPath }

    $networkPath = $env:HOMESHARE
    if (-not $networkPath) { Write-Verbose 'HOMESHARE not set; using Desktop'; return $desktopPath }
    if (!($networkPath -like "\\*")) { Write-Verbose 'Network path not UNC; using Desktop'; return $desktopPath }
    try {
        $result = Invoke-WithRetry -ScriptBlock {
            # Probe on a background runspace (in-process thread) so a dead server can be timed out. Start-Job was
            # used before, but it launches a new PowerShell process, which alone takes ~3-5 s in Windows PowerShell 5.1
            # - as long as the whole timeout - so a reachable share could be reported unreachable.
            $probe = [powershell]::Create().AddScript({
                param($Path)
                try {
                    if (Test-Path -LiteralPath $Path -PathType Container) {
                        $temp = Join-Path $Path "_bmtool_temp_$(Get-Random).txt"
                        [System.IO.File]::WriteAllText($temp, '')
                        [System.IO.File]::Delete($temp)
                        return $true
                    }
                } catch { }
                return $false
            }).AddArgument($networkPath)
            $isAccessible = $false; $timeout = $script:Config.NetworkTimeoutSeconds
            $handle = $probe.BeginInvoke()
            if ($handle.AsyncWaitHandle.WaitOne([int]($timeout * 1000))) {
                $isAccessible = [bool]($probe.EndInvoke($handle) | Select-Object -Last 1)
                $probe.Dispose()
            } else {
                # Don't wait for or dispose a probe stuck in network I/O; let it finish on its own thread
                Write-Log "Network path check timed out after $timeout s" 'DEBUG'
                $null = $probe.BeginStop($null, $null)
            }
            if (!$isAccessible) { throw "Network path not accessible: $networkPath" }
            $networkPath
        }
        Write-Verbose "Using network path: $result"; return $result
    } catch { Write-Verbose "Network path failed after retries: $_"; Write-Verbose "Using Desktop fallback: $desktopPath"; return $desktopPath }
}

# =====================================================================================
# BROWSER DETECTION & PROFILES
# =====================================================================================
function Test-BrowserRunning { param([Alias('BrowserName')][string]$Browser)
    $processMap = @{ 'chrome'=@('chrome','GoogleChromeHelper','Google Chrome Helper','Google Chrome Helper (Renderer)','Google Chrome Helper (GPU)','Google Chrome Helper (Plugin)','crashpad_handler'); 'msedge'=@('msedge','MicrosoftEdge','MicrosoftEdgeWebView2','msedgewebview2','MicrosoftEdgeCP','MicrosoftEdgeSH','identity_helper'); 'firefox'=@('firefox','plugin-container','firefox.exe','crashreporter','updater','maintenanceservice') }
    $key = $Browser.ToLower(); if ($key -eq 'edge') { $key = 'msedge' }
    if (-not $processMap.ContainsKey($key)) { Write-Warning "Unknown browser: $Browser"; return $false }
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
            # Profiles folder plus any profiles.ini entries stored elsewhere (custom locations)
            $dirs = @()
            $base = "$env:APPDATA\Mozilla\Firefox\Profiles"
            if (Test-Path $base) { $dirs += @(Get-ChildItem -Path $base -Directory) }
            $dirs += @(Get-FirefoxIniProfiles | Where-Object { Test-Path -LiteralPath $_.Path -PathType Container } | ForEach-Object { Get-Item -LiteralPath $_.Path })
            $dirs = @($dirs | Where-Object { Test-Path (Join-Path $_.FullName $FileName) } | Sort-Object FullName -Unique)
            foreach ($d in $dirs) { $profiles += @{ Name = $d.Name; Path = $d.FullName; LastUsed = $d.LastWriteTime } }
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

function Get-FirefoxIniProfiles {
    # Parses profiles.ini into profile entries, flagging the profile each Firefox install actually launches
    # ([Install*] Default=) and the legacy default ([Profile*] Default=1).
    $base = "$env:APPDATA\Mozilla\Firefox"
    $ini = Join-Path $base 'profiles.ini'
    if (!(Test-Path -LiteralPath $ini)) { Write-Verbose "Firefox profiles.ini not found: $ini"; return @() }

    $sections = [ordered]@{}; $current = $null
    foreach ($line in Get-Content -LiteralPath $ini) {
        $t = $line.Trim()
        if (!$t -or $t.StartsWith('#') -or $t.StartsWith(';')) { continue }
        if ($t -match '^\[(.+)\]$') { $current = $Matches[1]; $sections[$current] = @{}; continue }
        if ($current -and $t -match '^([^=]+)=(.*)$') { $sections[$current][$Matches[1].Trim()] = $Matches[2].Trim() }
    }

    function _Resolve([string]$p, [bool]$relative) {
        $p = $p -replace '/', '\'
        if ($relative -or -not [System.IO.Path]::IsPathRooted($p)) { Join-Path $base $p } else { $p }
    }

    $installDefaults = @(foreach ($k in $sections.Keys) {
        if ($k -like 'Install*' -and $sections[$k]['Default']) { [System.IO.Path]::GetFullPath((_Resolve $sections[$k]['Default'] $false)) }
    })

    foreach ($k in $sections.Keys) {
        if ($k -notlike 'Profile*') { continue }
        $s = $sections[$k]
        if (-not $s['Path']) { continue }
        $full = [System.IO.Path]::GetFullPath((_Resolve $s['Path'] ($s['IsRelative'] -eq '1')))
        [pscustomobject]@{
            Name           = Split-Path $full -Leaf
            Path           = $full
            InstallDefault = $installDefaults -contains $full
            LegacyDefault  = $s['Default'] -eq '1'
            HasPlaces      = Test-Path -LiteralPath (Join-Path $full 'places.sqlite')
        }
    }
}

function Get-FirefoxProfile {
    # Preference: the profile a Firefox install launches, then the legacy Default=1 profile, then the most
    # recently used profile with bookmarks. Only profiles that actually have places.sqlite are considered.
    $candidates = @(Get-FirefoxIniProfiles | Where-Object HasPlaces)
    $pick = $candidates | Where-Object InstallDefault |
        Sort-Object { (Get-Item -LiteralPath (Join-Path $_.Path 'places.sqlite')).LastWriteTime } -Descending | Select-Object -First 1
    if (-not $pick) { $pick = $candidates | Where-Object LegacyDefault | Select-Object -First 1 }
    if ($pick) { Write-Verbose "Found Firefox profile: $($pick.Path)"; return $pick.Path }

    Write-Verbose 'No default profile in profiles.ini has places.sqlite; using most recently used profile'
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
        if ($BrowserName -eq 'Firefox') { Copy-FirefoxPlaces -Src $sourceFile -Dst $backupPath } else { Copy-Item $sourceFile $backupPath -Force }
        Write-Log "Created backup: $backupPath"
        # Sort by the timestamp in the name: Copy-Item keeps the source's LastWriteTime, so file dates don't reflect backup order
        $backups = @(Get-ChildItem -LiteralPath $backupDir -Filter "$BookmarkFile.*.backup" | Sort-Object Name -Descending)
        if ($backups.Count -gt 10) { $backups | Select-Object -Skip 10 | ForEach-Object { Remove-Item -LiteralPath $_.FullName,"$($_.FullName)-wal","$($_.FullName)-shm" -Force -ErrorAction SilentlyContinue; Write-Verbose "Removed old backup: $($_.Name)" } }
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
                $fi = Get-Item -LiteralPath $FilePath
                if ($fi.Length -lt 100) { Write-Verbose 'Firefox file too small to be a SQLite database'; return $false }
                $header = New-Object byte[] 16
                $fs = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'ReadWrite')
                try { $null = $fs.Read($header, 0, 16) } finally { $fs.Dispose() }
                if (-not [System.Text.Encoding]::ASCII.GetString($header).StartsWith('SQLite format 3')) { Write-Verbose 'Not a SQLite database'; return $false }

                if (-not $script:SQLiteAvailable) { Write-Verbose "Firefox sqlite header valid (deep check skipped - SQLite unavailable): $FilePath"; return $true }

                # Deep check on a temp copy (never touches the source): must be a healthy Firefox places database
                $tmp = [System.IO.Path]::GetTempFileName(); $conn = $null; $cmd = $null
                try {
                    Copy-FirefoxPlaces -Src $FilePath -Dst $tmp
                    $conn = New-SQLiteConnection "Data Source=$tmp;Version=3;Pooling=False;"
                    $conn.Open()
                    $cmd = $conn.CreateCommand()
                    $cmd.CommandText = 'PRAGMA quick_check;'
                    $check = [string]$cmd.ExecuteScalar()
                    $cmd.CommandText = "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name IN ('moz_bookmarks','moz_places')"
                    $tables = [int]$cmd.ExecuteScalar()
                    $roots = 0
                    if ($tables -eq 2) { $cmd.CommandText = "SELECT COUNT(*) FROM moz_bookmarks WHERE guid IN ('root________','menu________','toolbar_____','unfiled_____')"; $roots = [int]$cmd.ExecuteScalar() }
                    $isValid = ($check -eq 'ok') -and ($tables -eq 2) -and ($roots -eq 4)
                    if ($isValid) { Write-Verbose "Firefox places database valid: $FilePath" } else { Write-Verbose "Invalid Firefox database (quick_check=$check, tables=$tables, roots=$roots)" }
                    return $isValid
                } finally {
                    if ($cmd) { $cmd.Dispose() }
                    if ($conn) { $conn.Dispose() }
                    Remove-Item -LiteralPath $tmp,"$tmp-wal","$tmp-shm" -Force -ErrorAction SilentlyContinue
                }
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
# Shared Netscape bookmark-file (the standard browser import/export HTML) helpers
$script:NetscapeHeader = @"
<!DOCTYPE NETSCAPE-Bookmark-file-1>
<!-- This is an automatically generated file.
     It will be read and overwritten.
     DO NOT EDIT! -->
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=UTF-8">
<TITLE>Bookmarks</TITLE>
<H1>Bookmarks</H1>

"@

function ConvertTo-HtmlText { param([AllowNull()][AllowEmptyString()][string]$Text) [System.Net.WebUtility]::HtmlEncode([string]$Text) }

function Get-NetscapeDateAttr {
    # Netscape format stores Unix time in seconds. Chrome uses microseconds since 1601; Firefox microseconds since 1970.
    param([string]$Name,$Value,[ValidateSet('Chrome','Firefox')][string]$Epoch)
    $n = [long]0
    if ($null -eq $Value -or -not [long]::TryParse([string]$Value, [ref]$n) -or $n -le 0) { return '' }
    $secs = [long][math]::Floor($n / 1000000)
    if ($Epoch -eq 'Chrome') { $secs -= 11644473600 }
    if ($secs -le 0) { return '' }
    " $Name=`"$secs`""
}

function Save-NetscapeHtml {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][System.Text.StringBuilder]$Body)
    $text = $script:NetscapeHeader + "<DL><p>`n" + $Body.ToString() + "</DL><p>`n"
    [System.IO.File]::WriteAllText($Path, $text, (New-Object System.Text.UTF8Encoding($false)))
}

function ConvertTo-ChromeHtml {
    param([Parameter(Mandatory)][string]$JsonPath,[Parameter(Mandatory)][string]$OutputPath)
    try {
        $json = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $sb = New-Object System.Text.StringBuilder

        function _Prop($obj, [string]$name) { $p = $obj.PSObject.Properties[$name]; if ($p) { $p.Value } else { $null } }

        function _Node($node, [int]$depth, [string]$extraAttr) {
            $pad = '    ' * $depth
            $type = _Prop $node 'type'
            $title = ConvertTo-HtmlText (_Prop $node 'name')
            $added = Get-NetscapeDateAttr 'ADD_DATE' (_Prop $node 'date_added') 'Chrome'
            if ($type -eq 'folder') {
                $modified = Get-NetscapeDateAttr 'LAST_MODIFIED' (_Prop $node 'date_modified') 'Chrome'
                $null = $sb.Append("$pad<DT><H3$added$modified$extraAttr>$title</H3>`n$pad<DL><p>`n")
                foreach ($child in @(_Prop $node 'children')) { if ($null -ne $child) { _Node $child ($depth + 1) '' } }
                $null = $sb.Append("$pad</DL><p>`n")
            } elseif ($type -eq 'url') {
                $href = ConvertTo-HtmlText (_Prop $node 'url')
                $null = $sb.Append("$pad<DT><A HREF=`"$href`"$added>$title</A>`n")
            }
        }

        $roots = _Prop $json 'roots'
        $bar = _Prop $roots 'bookmark_bar'
        if ($bar) { _Node $bar 1 ' PERSONAL_TOOLBAR_FOLDER="true"' }
        foreach ($rootName in 'other','synced') {
            $r = _Prop $roots $rootName
            if ($r -and @(_Prop $r 'children' | Where-Object { $null -ne $_ }).Count -gt 0) { _Node $r 1 '' }
        }

        Save-NetscapeHtml -Path $OutputPath -Body $sb
        Write-Log "Converted Chrome bookmarks to HTML: $OutputPath"
        return $true
    } catch {
        Write-Log "Failed to convert Chrome bookmarks to HTML: $_" 'ERROR'
        return $false
    }
}

function New-SQLiteConnection {
    param([Parameter(Mandatory)][string]$ConnectionString)
    # Try standard New-Object first
    try {
        return New-Object -TypeName System.Data.SQLite.SQLiteConnection -ArgumentList $ConnectionString -ErrorAction Stop
    } catch {
        # Fallback: Use Reflection to instantiate if the type isn't visible to PowerShell
        Write-Verbose "Standard instantiation failed, trying reflection..."
        $assembly = [System.AppDomain]::CurrentDomain.GetAssemblies() | Where-Object { $_.GetName().Name -eq 'System.Data.SQLite' } | Select-Object -First 1
        if ($null -eq $assembly) { throw "System.Data.SQLite assembly is not loaded." }

        $type = $assembly.GetType('System.Data.SQLite.SQLiteConnection')
        if ($null -eq $type) { throw "System.Data.SQLite.SQLiteConnection type not found in assembly." }

        return [Activator]::CreateInstance($type, @($ConnectionString))
    }
}

function Copy-FirefoxPlaces {
    # Copies places.sqlite plus its write-ahead log (places.sqlite-wal). While Firefox is running, recent
    # bookmark changes live only in the WAL, so copying the main file alone can miss them. The WAL is then
    # checkpointed into the copy so the destination is a single self-contained .sqlite file.
    param([Parameter(Mandatory)][string]$Src,[Parameter(Mandatory)][string]$Dst)

    $srcWal = "$Src-wal"; $dstWal = "$Dst-wal"; $dstShm = "$Dst-shm"
    Remove-Item -LiteralPath $dstWal,$dstShm -Force -ErrorAction SilentlyContinue
    Copy-Item -LiteralPath $Src -Destination $Dst -Force

    $hasWal = (Test-Path -LiteralPath $srcWal) -and (Get-Item -LiteralPath $srcWal).Length -gt 0
    # Header bytes 18/19 = 2 means the file is in WAL mode; opening it later (even read-only) would create
    # -wal/-shm files next to the backup, so it is switched to a standalone (rollback-journal) file below.
    $hdr = New-Object byte[] 20
    $fs = [System.IO.File]::Open($Dst, 'Open', 'Read', 'ReadWrite')
    try { $null = $fs.Read($hdr, 0, 20) } finally { $fs.Dispose() }
    $isWalMode = $hdr[18] -eq 2 -or $hdr[19] -eq 2
    if (-not $hasWal -and -not $isWalMode) { return }

    if ($hasWal) { Copy-Item -LiteralPath $srcWal -Destination $dstWal -Force }

    if (-not $script:SQLiteAvailable) {
        if ($hasWal) { Write-Log "System.Data.SQLite not available - kept Firefox WAL beside backup: $dstWal" 'WARN' }
        return
    }

    # Merge in a local temp copy, then copy the result over: SQLite cannot open UNC paths (\\server\share\...)
    $stage = [System.IO.Path]::GetTempFileName()
    $connection = $null; $command = $null
    try {
        Copy-Item -LiteralPath $Dst -Destination $stage -Force
        if ($hasWal) { Copy-Item -LiteralPath $srcWal -Destination "$stage-wal" -Force }
        $connection = New-SQLiteConnection "Data Source=$stage;Version=3;Pooling=False;"
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = 'PRAGMA wal_checkpoint(TRUNCATE);'
        $null = $command.ExecuteNonQuery()
        $command.CommandText = 'PRAGMA journal_mode=DELETE;'
        $null = $command.ExecuteNonQuery()
        # Undisposed commands keep the file handle open after Close()
        $command.Dispose(); $command = $null
        $connection.Close(); $connection.Dispose(); $connection = $null
        Copy-Item -LiteralPath $stage -Destination $Dst -Force
        Write-Verbose "Merged Firefox WAL into backup: $Dst"
    } catch {
        Write-Log "Failed to merge Firefox WAL into ${Dst}: $_ (WAL kept beside backup)" 'WARN'
        return
    } finally {
        if ($command) { $command.Dispose() }
        if ($connection) { $connection.Close(); $connection.Dispose() }
        Remove-Item -LiteralPath $stage,"$stage-wal","$stage-shm" -Force -ErrorAction SilentlyContinue
    }
    Remove-Item -LiteralPath $dstWal,$dstShm -Force -ErrorAction SilentlyContinue
}

function ConvertTo-FirefoxHtml {
    param([Parameter(Mandatory)][string]$SqlitePath,[Parameter(Mandatory)][string]$OutputPath)

    if (-not $script:SQLiteAvailable) {
        Write-Log "System.Data.SQLite not available - Firefox HTML conversion skipped" 'INFO'
        Write-Log "Firefox bookmarks exported as SQLite database successfully" 'INFO'
        return $false
    }

    $connection = $null
    $command = $null
    $reader = $null
    $tempDbPath = $null

    try {
        # Copy to temp file (with WAL, if any) to avoid UNC path issues and locks with SQLite
        $tempDbPath = [System.IO.Path]::GetTempFileName()
        Copy-FirefoxPlaces -Src $SqlitePath -Dst $tempDbPath

        $connection = New-SQLiteConnection "Data Source=$tempDbPath;Version=3;Read Only=True;Pooling=False;"
        $connection.Open()
        
        # Load the whole bookmark tree (type 1 = bookmark, 2 = folder, 3 = separator)
        $command = $connection.CreateCommand()
        $command.CommandText = "SELECT mb.id, mb.type, mb.parent, mb.title, mb.guid, mb.dateAdded, mb.lastModified, mp.url FROM moz_bookmarks mb LEFT JOIN moz_places mp ON mb.fk = mp.id ORDER BY mb.parent, mb.position"
        $reader = $command.ExecuteReader()
        $children = @{}; $byGuid = @{}
        function _Val($v) { if ($v -is [System.DBNull]) { $null } else { $v } }
        while ($reader.Read()) {
            $row = [pscustomobject]@{
                Id = [long]$reader['id']; Type = [int]$reader['type']; Parent = [long](_Val $reader['parent'])
                Title = [string](_Val $reader['title']); Guid = [string](_Val $reader['guid']); Url = [string](_Val $reader['url'])
                Added = _Val $reader['dateAdded']; Modified = _Val $reader['lastModified']
            }
            if (-not $children.ContainsKey($row.Parent)) { $children[$row.Parent] = New-Object System.Collections.Generic.List[object] }
            $children[$row.Parent].Add($row)
            if ($row.Guid) { $byGuid[$row.Guid] = $row }
        }
        $reader.Dispose(); $reader = $null

        $sb = New-Object System.Text.StringBuilder
        function _Kids([long]$id) { if ($children.ContainsKey($id)) { $children[$id] } else { @() } }
        function _Items([long]$parentId, [int]$depth) {
            $pad = '    ' * $depth
            foreach ($n in (_Kids $parentId)) {
                $added = Get-NetscapeDateAttr 'ADD_DATE' $n.Added 'Firefox'
                $modified = Get-NetscapeDateAttr 'LAST_MODIFIED' $n.Modified 'Firefox'
                switch ($n.Type) {
                    1 { if ($n.Url) { $null = $sb.Append("$pad<DT><A HREF=`"$(ConvertTo-HtmlText $n.Url)`"$added$modified>$(ConvertTo-HtmlText $n.Title)</A>`n") } }
                    2 { _Folder $n $depth '' $n.Title }
                    3 { $null = $sb.Append("$pad<HR>`n") }
                }
            }
        }
        function _Folder($n, [int]$depth, [string]$extraAttr, [string]$title) {
            $pad = '    ' * $depth
            $added = Get-NetscapeDateAttr 'ADD_DATE' $n.Added 'Firefox'
            $modified = Get-NetscapeDateAttr 'LAST_MODIFIED' $n.Modified 'Firefox'
            $null = $sb.Append("$pad<DT><H3$added$modified$extraAttr>$(ConvertTo-HtmlText $title)</H3>`n$pad<DL><p>`n")
            _Items $n.Id ($depth + 1)
            $null = $sb.Append("$pad</DL><p>`n")
        }

        # Same layout as Firefox's own HTML export: menu items at top level, then the special folders.
        # The tags root is skipped (tags are not bookmarks).
        if ($byGuid.ContainsKey('menu________')) { _Items $byGuid['menu________'].Id 1 }
        if ($byGuid.ContainsKey('toolbar_____')) { _Folder $byGuid['toolbar_____'] 1 ' PERSONAL_TOOLBAR_FOLDER="true"' 'Bookmarks Toolbar' }
        if ($byGuid.ContainsKey('unfiled_____') -and @(_Kids $byGuid['unfiled_____'].Id).Count -gt 0) { _Folder $byGuid['unfiled_____'] 1 ' UNFILED_BOOKMARKS_FOLDER="true"' 'Other Bookmarks' }
        if ($byGuid.ContainsKey('mobile______') -and @(_Kids $byGuid['mobile______'].Id).Count -gt 0) { _Folder $byGuid['mobile______'] 1 '' 'Mobile Bookmarks' }

        $connection.Close()
        Save-NetscapeHtml -Path $OutputPath -Body $sb
        Write-Log "Converted Firefox bookmarks to HTML: $OutputPath"
        return $true
    } catch {
        Write-Log "Failed to convert Firefox bookmarks to HTML: $_" 'WARN'
        if ($connection -and $connection.State -eq 'Open') { $connection.Close() }
        return $false
    } finally {
        # Undisposed readers/commands keep the temp file locked, so it couldn't be deleted
        if ($reader) { $reader.Dispose() }
        if ($command) { $command.Dispose() }
        if ($connection) { $connection.Dispose() }
        if ($tempDbPath) {
            Remove-Item -LiteralPath $tempDbPath,"$tempDbPath-wal","$tempDbPath-shm" -Force -ErrorAction SilentlyContinue
        }
    }
}

# =====================================================================================
# ZIP ARCHIVE FUNCTIONS
# =====================================================================================
function New-ZipArchive {
    param([Parameter(Mandatory)][string]$SourcePath,[Parameter(Mandatory)][string]$ZipPath,
          # Only include files whose name contains this (e.g. the run's timestamp), so older exports in the folder are left out
          [string]$NameContains)
    try {
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $files = @(Get-ChildItem -LiteralPath $SourcePath -File | Where-Object { $_.Extension -in '.json','.sqlite','.html','.htm' -and (-not $NameContains -or $_.Name.Contains($NameContains)) })
        
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
        [switch]$ExportHtmlOnly,
        [switch]$AllProfiles,
        [switch]$CreateZip
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
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\s','_')" } else { '' }
                
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
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\s','_')" } else { '' }
                
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
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\s','_')" } else { '' }
                
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
                    if ($PSCmdlet.ShouldProcess($dstSqlite, "Export Firefox bookmarks from $src")) {
                        Copy-FirefoxPlaces -Src $src -Dst $dstSqlite
                        Write-Log "SUCCESS: Exported Firefox from $src to $dstSqlite"
                        $script:OperationResults += [pscustomobject]@{ Browser='Firefox'; Success=$true; Message="Exported to $dstSqlite" }
                    }
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
        if (New-ZipArchive -SourcePath $Path -ZipPath $zipPath -NameContains "_$timestamp") {
            $script:OperationResults += [pscustomobject]@{ Browser='ZIP'; Success=$true; Message="Created archive: $zipPath" }
        } else {
            $script:OperationResults += [pscustomobject]@{ Browser='ZIP'; Success=$false; Message="ZIP creation failed" }
        }
    }
}

function Find-BookmarkImportSource {
    # Locates the file to import for a browser in $Path. Accepts this tool's export names
    # (<Browser>[-<Profile>]_BookmarkData_<timestamp><ext>, newest wins) and the legacy names
    # (<LegacyBase>[-<Profile>]<ext>). Returns $null if nothing matches.
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Browser,
        [Parameter(Mandatory)][string]$Extension,
        [Parameter(Mandatory)][string]$LegacyBase,
        [string]$ProfileSuffix,
        # Raw profile name: older exports kept spaces in the file name ("Chrome-Profile 1_...")
        [string]$ProfileName
    )

    function _Newest([string]$Filter) {
        Get-ChildItem -LiteralPath $Path -Filter $Filter -File -ErrorAction SilentlyContinue |
            Sort-Object @{ Expression = { if ($_.Name -match '_BookmarkData_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2})') { $Matches[1] } else { '' } } }, LastWriteTime -Descending |
            Select-Object -First 1 -ExpandProperty FullName
    }
    function _Exact([string]$Name) { $f = Join-Path $Path $Name; if (Test-Path -LiteralPath $f -PathType Leaf) { $f } }

    $found = $null
    $suffixes = @(@($ProfileSuffix, $(if ($ProfileName) { "-$ProfileName" })) | Where-Object { $_ } | Select-Object -Unique)
    foreach ($sfx in $suffixes) {
        if (-not $found) { $found = _Newest "$Browser$sfx`_BookmarkData_*$Extension" }
        if (-not $found) { $found = _Exact "$LegacyBase$sfx$Extension" }
    }
    if ($suffixes) {
        # Importing a specific profile (-AllProfiles): only that profile's own export may be used. Falling back
        # to another profile's file would silently replace this profile's bookmarks with someone else's.
        if ($found) { Write-Log "Selected $Browser import file: $found" } else { Write-Log "No $Browser export found for profile '$($suffixes[0].TrimStart('-'))' in $Path" 'WARN' }
        return $found
    }
    if (-not $found) { $found = _Newest "$Browser`_BookmarkData_*$Extension" }
    if (-not $found) { $found = _Exact "$LegacyBase$Extension" }
    # Last resort: any per-profile export for this browser (e.g. from an -AllProfiles export or the desktop app)
    if (-not $found) { $found = _Newest "$Browser*_BookmarkData_*$Extension" }

    if ($found) { Write-Log "Selected $Browser import file: $found" }
    $found
}

function Import-Bookmarks {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Path,[switch]$Chrome,[switch]$Edge,[switch]$Firefox,[switch]$CloseBrowserIfRunning,[switch]$AllProfiles)

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
        if ($script:Config.VerifyFileIntegrity -ne $false -and -not (Test-BookmarkFileIntegrity -FilePath $Src -BrowserType $Label)) {
            Write-Log "$Label import skipped - $Src is not a valid $Label bookmark file (profile left unchanged)" 'ERROR'
            $script:OperationResults += [pscustomobject]@{ Browser=$Label; Success=$false; Message="Invalid bookmark file: $Src" }
            return
        }
        if ($script:Config.AutoBackupBeforeImport) { $null = Backup-ExistingBookmarks -BrowserProfile (Split-Path $Dst -Parent) -BookmarkFile (Split-Path $Dst -Leaf) -BrowserName $Label }
        if ($PSCmdlet.ShouldProcess($Dst, "Import $Label bookmarks from $Src")) {
            if ($Label -eq 'Firefox') {
                # A leftover WAL/SHM (e.g. after Firefox was force-closed) would be replayed on top of the
                # imported database at next launch, undoing or corrupting the import. Its contents are already
                # in the pre-import backup above.
                foreach ($side in "$Dst-wal","$Dst-shm") {
                    if (Test-Path -LiteralPath $side) {
                        try { Remove-Item -LiteralPath $side -Force }
                        catch { throw "Cannot remove $side - is Firefox still running? Close Firefox and retry the import. ($_)" }
                        Write-Verbose "Removed stale Firefox file: $side"
                    }
                }
            }
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
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\s','_')" } else { '' }
                $src = Find-BookmarkImportSource -Path $Path -Browser 'Chrome' -Extension '.json' -LegacyBase 'Chrome-Bookmarks' -ProfileSuffix $suffix -ProfileName $(if ($AllProfiles) { $pr.Name })
                if (-not $src) { $src = Join-Path $Path "Chrome$suffix`_BookmarkData_*.json" }
                if (!(Test-Path -LiteralPath $src) -and -not $Silent) { Write-Log "Prompting for Chrome file" 'INFO'; $tmp = Select-FileDialog -Filter 'Chrome Bookmarks (*.json)|*.json'; if ($tmp) { $src = $tmp; Write-Log "User selected Chrome import file: $src" } }
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
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\s','_')" } else { '' }
                $src = Find-BookmarkImportSource -Path $Path -Browser 'Edge' -Extension '.json' -LegacyBase 'Edge-Bookmarks' -ProfileSuffix $suffix -ProfileName $(if ($AllProfiles) { $pr.Name })
                if (-not $src) { $src = Join-Path $Path "Edge$suffix`_BookmarkData_*.json" }
                if (!(Test-Path -LiteralPath $src) -and -not $Silent) { Write-Log "Prompting for Edge file" 'INFO'; $tmp = Select-FileDialog -Filter 'Edge Bookmarks (*.json)|*.json'; if ($tmp) { $src = $tmp; Write-Log "User selected Edge import file: $src" } }
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
                $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\s','_')" } else { '' }
                $src = Find-BookmarkImportSource -Path $Path -Browser 'Firefox' -Extension '.sqlite' -LegacyBase 'Firefox-places' -ProfileSuffix $suffix -ProfileName $(if ($AllProfiles) { $pr.Name })
                if (-not $src) { $src = Join-Path $Path "Firefox$suffix`_BookmarkData_*.sqlite" }
                if (!(Test-Path -LiteralPath $src) -and -not $Silent) { Write-Log "Prompting for Firefox file" 'INFO'; $tmp = Select-FileDialog -Filter 'Firefox bookmark database (*.sqlite)|*.sqlite'; if ($tmp) { $src = $tmp; Write-Log "User selected Firefox import file: $src" } }
                _CopyIn -Src $src -Dst $dst -Label 'Firefox'
            }
        } else { Write-Log 'Firefox profile(s) not found' 'WARN' }
    }

    Write-Log 'Import operation completed'
}

function Import-FromZip {
    param([Parameter(Mandatory)][string]$ZipPath,[string]$TargetBrowser,[switch]$AllProfiles)
    
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
            Import-Bookmarks -Path $tempExtractPath -Chrome -AllProfiles:$AllProfiles
            $importSuccess = $true
        } elseif ($TargetBrowser -eq 'Edge' -and $edgeFiles) {
            Write-Log "Importing Edge bookmarks from ZIP"
            Import-Bookmarks -Path $tempExtractPath -Edge -AllProfiles:$AllProfiles
            $importSuccess = $true
        } elseif ($TargetBrowser -eq 'Firefox' -and $firefoxFiles) {
            Write-Log "Importing Firefox bookmarks from ZIP"
            Import-Bookmarks -Path $tempExtractPath -Firefox -AllProfiles:$AllProfiles
            $importSuccess = $true
        } else {
            # Auto-detect and import all available
            Write-Log "Auto-detecting browsers in ZIP archive"
            if ($chromeFiles) { Import-Bookmarks -Path $tempExtractPath -Chrome -AllProfiles:$AllProfiles; $importSuccess = $true }
            if ($edgeFiles) { Import-Bookmarks -Path $tempExtractPath -Edge -AllProfiles:$AllProfiles; $importSuccess = $true }
            if ($firefoxFiles) { Import-Bookmarks -Path $tempExtractPath -Firefox -AllProfiles:$AllProfiles; $importSuccess = $true }
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
    $form.Text = "Bookmark Backup Tool v$script:ToolVersion - Enhanced Edition"; $form.Size = '600,480'; $form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false

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
        Export-Bookmarks -Path $currentPath -Chrome:$chkChrome.Checked -Edge:$chkEdge.Checked -Firefox:$chkFirefox.Checked -ExportHtmlOnly:$chkHtmlOnly.Checked -AllProfiles:$chkAllProfiles.Checked -CreateZip:$chkCreateZip.Checked
        [Windows.Forms.MessageBox]::Show('Export completed. Check the log for details.','Export Complete',[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information) | Out-Null
        Write-Log 'Export completed via GUI'
    })
    $form.Controls.Add($btnExport)

    $btnImport = New-Object Windows.Forms.Button; $btnImport.Text='Import Bookmarks'; $btnImport.Size='180,40'; $btnImport.Location='260,300'
    $btnImport.Add_Click({
        Write-Log 'Import button clicked'
        $currentPath = if ($txtPath.Text) { $txtPath.Text } else { Get-HomeSharePath }
        $txtPath.Text = $currentPath
        Import-Bookmarks -Path $currentPath -Chrome:$chkChrome.Checked -Edge:$chkEdge.Checked -Firefox:$chkFirefox.Checked -CloseBrowserIfRunning -AllProfiles:$chkAllProfiles.Checked
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
            'daily'  { $trigger = New-ScheduledTaskTrigger -Daily -At $Time }
            'weekly' { $trigger = New-ScheduledTaskTrigger -Weekly -At $Time -DaysOfWeek Sunday }
            'monthly' {
                # New-ScheduledTaskTrigger has no -Monthly parameter. Build the supported
                # Task Scheduler monthly trigger as a client-only CIM instance instead.
                $start = [datetime]::Today.Add([timespan]::Parse($Time))
                if ($start -le (Get-Date)) { $start = $start.AddMonths(1) }
                $trigger = New-CimInstance -ClassName MSFT_TaskMonthlyTrigger `
                    -Namespace Root/Microsoft/Windows/TaskScheduler `
                    -ClientOnly `
                    -Property @{
                        Enabled       = $true
                        StartBoundary = $start.ToString('s')
                        DaysOfMonth   = [uint32]1
                        MonthsOfYear  = [uint16]4095
                    }
            }
            default { throw "Invalid frequency: $Frequency. Must be Daily, Weekly, or Monthly." }
        }
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
        $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive

        $taskName = 'BookmarkBackupTool_AutoExport'
        $existing = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($existing) { Write-Log "Scheduled task exists: $taskName"; Unregister-ScheduledTask -TaskName $taskName -Confirm:$false; Write-Log "Removed existing scheduled task: $taskName" }
        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Automatic bookmark backup using BookmarkTool v$script:ToolVersion" | Out-Null
        Write-Log "Successfully created scheduled task: $taskName"; Write-Log "Frequency: $Frequency at $Time"; Write-Information "[SUCCESS] Scheduled task created. Next Run: $($trigger.StartBoundary)"; $true
    } catch { Write-Error "Failed to create scheduled task: $_"; Write-Log "ERROR: Failed to create scheduled task - $_" 'ERROR'; $false }
}

function Remove-BookmarkScheduledTask { try { $name='BookmarkBackupTool_AutoExport'; if (Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue) { Unregister-ScheduledTask -TaskName $name -Confirm:$false; Write-Log "Removed scheduled task: $name"; $true } else { Write-Warning "Scheduled task not found: $name"; $false } } catch { Write-Error "Failed to remove scheduled task: $_"; Write-Log "ERROR: Failed to remove scheduled task - $_" 'ERROR'; $false } }

# =====================================================================================
# PUBLIC HELPER COMMANDS
# (exported by the module; also usable after dot-sourcing the script: . .\BookMarkToolv5.ps1)
# =====================================================================================
function Get-BookmarkConfiguration {
    <# .SYNOPSIS Returns the current configuration (defaults merged with BookmarkTool.config.json). #>
    [CmdletBinding()]
    param([string]$ConfigFilePath)
    if ($ConfigFilePath) { Get-Configuration -ConfigFilePath $ConfigFilePath } else { $script:Config }
}

function Set-BookmarkConfiguration {
    <#
    .SYNOPSIS Changes configuration settings and saves them to BookmarkTool.config.json.
    .EXAMPLE Set-BookmarkConfiguration -AutoBackupBeforeImport $false -LogRetentionDays 60
    .EXAMPLE $c = Get-BookmarkConfiguration; $c.DefaultPath = 'D:\Backups'; Set-BookmarkConfiguration -Configuration $c
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [hashtable]$Configuration,
        [string]$DefaultPath,
        [bool]$PreferNetworkPath,
        [ValidateSet('Chrome','Edge','Firefox')][string[]]$DefaultBrowsers,
        [bool]$AutoBackupBeforeImport,
        [bool]$VerifyFileIntegrity,
        [ValidateRange(1,300)][int]$NetworkTimeoutSeconds,
        [ValidateRange(1,20)][int]$MaxRetryAttempts,
        [ValidateRange(0,60)][int]$RetryDelaySeconds,
        [ValidateRange(1,3650)][int]$LogRetentionDays,
        [bool]$DetailedLogging,
        [string]$ConfigFilePath
    )
    $cfg = if ($Configuration) { $Configuration.Clone() } else { $script:Config.Clone() }
    foreach ($k in $PSBoundParameters.Keys) {
        if ($k -in 'Configuration','ConfigFilePath','WhatIf','Confirm','Verbose','Debug','ErrorAction','WarningAction','InformationAction','ErrorVariable','WarningVariable','InformationVariable','OutVariable','OutBuffer','PipelineVariable') { continue }
        $cfg[$k] = $PSBoundParameters[$k]
    }
    $target = if ($ConfigFilePath) { $ConfigFilePath } else { Join-Path $env:USERPROFILE 'BookmarkTool.config.json' }
    if ($PSCmdlet.ShouldProcess($target, 'Save bookmark tool configuration')) {
        if (Save-Configuration -Config $cfg -ConfigFilePath $target) { $script:Config = $cfg; $script:HomeSharePathCache = $null; $cfg }
    }
}

function Test-BrowserInstalled {
    <# .SYNOPSIS Returns $true if the browser has a user profile folder on this computer. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][Alias('Browser')][ValidateSet('Chrome','Edge','Firefox')][string]$BrowserName)
    $dir = switch ($BrowserName) {
        'Chrome'  { "$env:LOCALAPPDATA\Google\Chrome\User Data" }
        'Edge'    { "$env:LOCALAPPDATA\Microsoft\Edge\User Data" }
        'Firefox' { "$env:APPDATA\Mozilla\Firefox" }
    }
    Test-Path -LiteralPath $dir -PathType Container
}

function Get-BrowserProfiles {
    <# .SYNOPSIS Lists browser profiles that contain bookmarks: the one the tool uses by default, or all with -AllProfiles. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][Alias('Browser')][ValidateSet('Chrome','Edge','Firefox')][string]$BrowserName, [switch]$AllProfiles)
    $file = if ($BrowserName -eq 'Firefox') { 'places.sqlite' } else { 'Bookmarks' }
    $paths = if ($AllProfiles) { @(Get-AllBrowserProfiles -Browser $BrowserName -FileName $file | ForEach-Object { $_.Path }) }
             else { @(switch ($BrowserName) { 'Chrome' { Get-ChromeProfile } 'Edge' { Get-EdgeProfile } 'Firefox' { Get-FirefoxProfile } }) }
    foreach ($p in $paths | Where-Object { $_ }) {
        $item = Get-Item -LiteralPath $p
        [pscustomobject]@{ Browser = $BrowserName; Name = $item.Name; FullName = $item.FullName; LastWriteTime = (Get-Item -LiteralPath (Join-Path $p $file)).LastWriteTime }
    }
}

function Test-BookmarkPrerequisites {
    <# .SYNOPSIS Checks PowerShell version, .NET assemblies and (optionally) write access to a target folder. #>
    [CmdletBinding()]
    param([string]$TargetPath)
    Test-Prerequisites -TargetPath $TargetPath
}

# =====================================================================================
# MAIN EXECUTION LOGIC
# (a function, so the PowerShell module built from this script by Build-Module.ps1 can expose it)
# =====================================================================================
function Invoke-BookmarkBackupTool {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [switch]$Silent,
        [ValidateSet('Export','Import')][string]$Action,
        [switch]$Chrome,
        [switch]$Edge,
        [switch]$Firefox,
        [ValidateScript({ if ($_ -and !(Test-Path $_ -IsValid)) { throw "Invalid path format: $_" }; $true })][string]$TargetPath,
        [switch]$HtmlOnly,
        [switch]$AllProfiles,
        [switch]$CreateZip,
        [switch]$Force,
        [switch]$CreateScheduledTask,
        [ValidateSet('Daily','Weekly','Monthly')][string]$ScheduleFrequency = 'Daily',
        [string]$ConfigPath
    )

    # A new run: forget any share-detection result cached by a previous call in the same session (module use)
    $script:HomeSharePathCache = $null
    if ($ConfigPath) { $script:Config = Get-Configuration -ConfigFilePath $ConfigPath }

    Invoke-LogRetention
    Write-Log "=== Bookmark Backup Tool v$script:ToolVersion Enhanced Edition Started ==="
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
            if ($script:IsModule) {
                Write-Warning 'The GUI needs an STA thread. If it fails, restart PowerShell with -STA (Windows PowerShell 5.1 is STA by default).'
            } else {
                Write-Verbose 'Re-launching in STA mode for Windows Forms...'
                $argsJoined = ($PSBoundParameters.GetEnumerator() | ForEach-Object {
                    if ($_.Value -is [switch]) { if ($_.Value.IsPresent) { "-$($_.Key)" } }
                    elseif ($_.Value -is [string]) { "-$($_.Key) `"$($_.Value)`"" }
                    else { "-$($_.Key) $($_.Value)" }
                }) -join ' '
                Start-Process -FilePath (Get-Process -Id $PID).Path -ArgumentList "-STA -NoProfile -ExecutionPolicy Bypass -File `"$script:ScriptPath`" $argsJoined" | Out-Null
                return
            }
        }

        if ($CreateScheduledTask) {
            Write-Log 'Creating scheduled task for automatic backups'
            if (New-BookmarkScheduledTask -Frequency $ScheduleFrequency) { Write-Log 'Scheduled task creation completed successfully'; return } else { throw 'Scheduled task creation failed.' }
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
                    if ($PSCmdlet.ShouldProcess($finalPath,'Export bookmarks')) { Export-Bookmarks -Path $finalPath -Chrome:$Chrome -Edge:$Edge -Firefox:$Firefox -ExportHtmlOnly:$HtmlOnly -AllProfiles:$AllProfiles -CreateZip:$CreateZip }
                    $summary = New-OperationSummary -Operations $script:OperationResults -StartTime $operationStart -OperationType 'Export'
                    Write-Log $summary
                }
                'Import' {
                    Write-Log "Silent import: Chrome=$Chrome Edge=$Edge Firefox=$Firefox; Path=$finalPath"
                    if ($PSCmdlet.ShouldProcess($finalPath,'Import bookmarks')) { Import-Bookmarks -Path $finalPath -Chrome:$Chrome -Edge:$Edge -Firefox:$Firefox -AllProfiles:$AllProfiles -CloseBrowserIfRunning:$Force }
                    $summary = New-OperationSummary -Operations $script:OperationResults -StartTime $operationStart -OperationType 'Import'
                    Write-Log $summary
                }
                default { throw "When using -Silent, -Action must be 'Export' or 'Import'." }
            }
            Write-Log '=== Silent mode execution completed successfully ==='
        } else {
            Write-Log 'Launching enhanced GUI mode'
            Show-GUI
            Write-Log '=== GUI mode execution completed successfully ==='
        }
    }
    catch {
        Write-Log "ERROR: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

# >>> SCRIPT ENTRY POINT - removed by Build-Module.ps1 when building the module
if ($MyInvocation.InvocationName -ne '.' -and $MyInvocation.Line -notmatch '^\s*\.\s') {
    try { Invoke-BookmarkBackupTool @PSBoundParameters; $global:LASTEXITCODE = 0; exit 0 }
    catch { $global:LASTEXITCODE = 1; exit 1 }
}
# <<< SCRIPT ENTRY POINT