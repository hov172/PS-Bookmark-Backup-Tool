# =====================================================================================
# Bookmark Backup Tool v5.0 - Enhanced Edition (Best-Practices Consolidated)
# Author: Jesus M. Ayala
# Version: 5.0
# Last Modified: Nov 11th, 2025
# Requires: PowerShell 5.1+, Windows 10/11
# License: MIT
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
    [string]$ConfigPath,

    [Parameter(HelpMessage="Process all browser profiles instead of just the latest")]
    [switch]$AllProfiles
)

# Hardening & sane defaults
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'

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
                try { if (Test-Path -Path $Path -PathType Container) { $temp = Join-Path $Path "_bmtool_temp_$(Get-Random).txt"; New-Item -Path $temp -ItemType File -Force | Out-Null; Remove-Item -Path $temp -Force | Out-Null; return $true } } catch { }
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
    $successful = $Operations | Where-Object { $_.Success }
    $failed     = $Operations | Where-Object { -not $_.Success }
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
# EXPORT / IMPORT
# =====================================================================================
function Export-Bookmarks {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Path,[switch]$Chrome,[switch]$Edge,[switch]$Firefox)

    if (!(Test-Path -LiteralPath $Path -PathType Container)) {
        if ($PSCmdlet.ShouldProcess($Path,'Create backup directory')) { New-Item -ItemType Directory -Path $Path -Force | Out-Null; Write-Log "Created backup directory: $Path" }
    }

    function _Copy { param([string]$Src,[string]$Dst,[string]$Label)
        if ($PSCmdlet.ShouldProcess($Dst, "Export $Label bookmarks from $Src")) {
            Copy-Item -LiteralPath $Src -Destination $Dst -Force
            Write-Log "SUCCESS: Exported $Label from $Src to $Dst"
            $script:OperationResults += [pscustomobject]@{ Browser=$Label; Success=$true; Message="Exported to $Dst" }
        }
    }

    if ($Chrome) {
        Write-Log 'Starting Chrome export...'
        $profiles = if ($AllProfiles) { Get-AllBrowserProfiles -Browser 'Chrome' -FileName 'Bookmarks' } else { $p = Get-ChromeProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($profiles) { foreach ($pr in $profiles) { $src = Join-Path $pr.Path 'Bookmarks'; $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\\s','_')" } else { '' }; $dst = Join-Path $Path ("Chrome-Bookmarks$suffix.json"); _Copy -Src $src -Dst $dst -Label 'Chrome' } }
        else { Write-Log 'Chrome profile(s) not found - skipping Chrome export' 'WARN'; $script:OperationResults += [pscustomobject]@{ Browser='Chrome'; Success=$false; Message='Profile not found' } }
    }

    if ($Edge) {
        Write-Log 'Starting Edge export...'
        $profiles = if ($AllProfiles) { Get-AllBrowserProfiles -Browser 'Edge' -FileName 'Bookmarks' } else { $p = Get-EdgeProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($profiles) { foreach ($pr in $profiles) { $src = Join-Path $pr.Path 'Bookmarks'; $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\\s','_')" } else { '' }; $dst = Join-Path $Path ("Edge-Bookmarks$suffix.json"); _Copy -Src $src -Dst $dst -Label 'Edge' } }
        else { Write-Log 'Edge profile(s) not found - skipping Edge export' 'WARN'; $script:OperationResults += [pscustomobject]@{ Browser='Edge'; Success=$false; Message='Profile not found' } }
    }

    if ($Firefox) {
        Write-Log 'Starting Firefox export...'
        $profiles = if ($AllProfiles) { Get-AllBrowserProfiles -Browser 'Firefox' -FileName 'places.sqlite' } else { $p = Get-FirefoxProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($profiles) { foreach ($pr in $profiles) { $src = Join-Path $pr.Path 'places.sqlite'; $suffix = if ($AllProfiles) { "-$($pr.Name -replace '\\s','_')" } else { '' }; $dst = Join-Path $Path ("Firefox-places$suffix.sqlite"); _Copy -Src $src -Dst $dst -Label 'Firefox' } }
        else { Write-Log 'Firefox profile(s) not found - skipping Firefox export' 'WARN'; $script:OperationResults += [pscustomobject]@{ Browser='Firefox'; Success=$false; Message='Profile not found' } }
    }

    Write-Log 'Export operation completed'
}

function Import-Bookmarks {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][string]$Path,[switch]$Chrome,[switch]$Edge,[switch]$Firefox)

    if ($Chrome  -and (Test-BrowserRunning 'chrome'))  { Write-Log 'ABORT: Chrome running'  'WARN'; return }
    if ($Edge    -and (Test-BrowserRunning 'msedge'))  { Write-Log 'ABORT: Edge running'    'WARN'; return }
    if ($Firefox -and (Test-BrowserRunning 'firefox')) { Write-Log 'ABORT: Firefox running' 'WARN'; return }

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
        $targets = if ($AllProfiles) { Get-AllBrowserProfiles -Browser 'Chrome' -FileName 'Bookmarks' } else { $p = Get-ChromeProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($targets) {
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
        $targets = if ($AllProfiles) { Get-AllBrowserProfiles -Browser 'Edge' -FileName 'Bookmarks' } else { $p = Get-EdgeProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($targets) {
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
        $targets = if ($AllProfiles) { Get-AllBrowserProfiles -Browser 'Firefox' -FileName 'places.sqlite' } else { $p = Get-FirefoxProfile; $arr=@(); if ($p){$arr+=@{Name='DefaultOrLatest';Path=$p}}; $arr }
        if ($targets) {
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
    $form.Text = 'Bookmark Backup Tool v5.0 - Enhanced Edition'; $form.Size = '600,420'; $form.StartPosition = 'CenterScreen'; $form.FormBorderStyle = 'FixedDialog'; $form.MaximizeBox = $false

    $chkChrome  = New-Object Windows.Forms.CheckBox; $chkChrome.Text='Chrome';  $chkChrome.Location='30,30';  $chkChrome.AutoSize=$true; $form.Controls.Add($chkChrome)
    $chkEdge    = New-Object Windows.Forms.CheckBox; $chkEdge.Text='Edge';      $chkEdge.Location='30,60';  $chkEdge.AutoSize=$true; $form.Controls.Add($chkEdge)
    $chkFirefox = New-Object Windows.Forms.CheckBox; $chkFirefox.Text='Firefox';$chkFirefox.Location='30,90'; $chkFirefox.AutoSize=$true; $form.Controls.Add($chkFirefox)

    $lblPath = New-Object Windows.Forms.Label; $lblPath.Text='Save/Load Path:'; $lblPath.Location='30,130'; $lblPath.Size='420,20'; $form.Controls.Add($lblPath)
    $txtPath = New-Object Windows.Forms.TextBox; $txtPath.Text = Get-HomeSharePath; $txtPath.Location='30,155'; $txtPath.Size='360,25'; $form.Controls.Add($txtPath)
    $btnBrowse = New-Object Windows.Forms.Button; $btnBrowse.Text='Browse'; $btnBrowse.Location='400,154'; $btnBrowse.Size='60,25'; $btnBrowse.Add_Click({ $dialog = New-Object Windows.Forms.FolderBrowserDialog; if ($dialog.ShowDialog() -eq 'OK') { $txtPath.Text = $dialog.SelectedPath; Write-Log "User selected path: $($dialog.SelectedPath)" } }); $form.Controls.Add($btnBrowse)

    $btnExport = New-Object Windows.Forms.Button; $btnExport.Text='Export Bookmarks'; $btnExport.Size='180,40'; $btnExport.Location='40,220'; $btnExport.Add_Click({ Write-Log 'Export button clicked'; $currentPath = if ($txtPath.Text) { $txtPath.Text } else { Get-HomeSharePath }; $txtPath.Text = $currentPath; Export-Bookmarks -Path $currentPath -Chrome:$chkChrome.Checked -Edge:$chkEdge.Checked -Firefox:$chkFirefox.Checked; [Windows.Forms.MessageBox]::Show('Export completed. Check the log for details.','Export Complete',[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information) | Out-Null; Write-Log 'Export completed via GUI' }); $form.Controls.Add($btnExport)

    $btnImport = New-Object Windows.Forms.Button; $btnImport.Text='Import Bookmarks'; $btnImport.Size='180,40'; $btnImport.Location='260,220'; $btnImport.Add_Click({ Write-Log 'Import button clicked'; $currentPath = if ($txtPath.Text) { $txtPath.Text } else { Get-HomeSharePath }; $txtPath.Text = $currentPath; Import-Bookmarks -Path $currentPath -Chrome:$chkChrome.Checked -Edge:$chkEdge.Checked -Firefox:$chkFirefox.Checked; [Windows.Forms.MessageBox]::Show('Import completed. Check the log for details.','Import Complete',[Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information) | Out-Null; Write-Log 'Import completed via GUI' }); $form.Controls.Add($btnImport)

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
                    Write-Log "Silent export: Chrome=$Chrome Edge=$Edge Firefox=$Firefox; Path=$finalPath"
                    if ($PSCmdlet.ShouldProcess($finalPath,'Export bookmarks')) { Export-Bookmarks -Path $finalPath -Chrome:$Chrome -Edge:$Edge -Firefox:$Firefox }
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
