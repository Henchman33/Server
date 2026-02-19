# ServerTempCleanup_SCCM.ps1
<# Save this as ServerTempCleanup_SCCM.ps1. Designed to run as SYSTEM in 64‑bit PowerShell (as SCCM does for packages/programs).
Note: -WhatIf supported — use it for pilots.

Pilot first (SCCM Run Script or Package/Program)
Run Script (fastest pilot)

Software Library → Scripts → Create Script → paste the script above.
Approve the script and Run it against a small pilot collection.
Use these parameters for a safe preview:
PowerShell-WhatIf -ListFiles -MinimumAgeHours 0 -CcmCacheMaxAgeDays 32 -CleanCcmCache $true -CleanOldProfiles $trueShow more lines

Check logs on each pilot:
C:\TempCleanup\Logs\TempCleanup_<SERVER>_<YYYYMMDD>.log

Package/Program (pilot or prod)

Command line (64‑bit PowerShell, no profile):
PowerShell%WinDir%\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File ServerTempCleanup_SCCM.ps1 -WhatIf -ListFiles -MinimumAgeHours 0 -CcmCacheMaxAgeDays 32 -CleanCcmCache $true -CleanOldProfiles $true -SilentShow more lines

Program settings:

Run: Whether or not a user is logged on
Run with Administrative rights
Do not allow users to interact
Never rerun (for one‑time push)
Timeout: 60 minutes

Important: Ensure SCCM uses 64‑bit PowerShell. For packages/programs the System32 path above ensures 64‑bit. For Run Scripts, enable the client setting: “Run script execution in 64-bit PowerShell (when available).”

🚀 Production run (no WhatIf)
When you’re satisfied with the preview metrics, deploy without -WhatIf. Example:
PowerShell%WinDir%\System32\WindowsPowerShell\v1.0\powershell.exe -ExecutionPolicy Bypass -NoProfile -NonInteractive -File ServerTempCleanup_SCCM.ps1 -MinimumAgeHours 48 -IncludeLogs -CcmCacheMaxAgeDays 32 -CleanCcmCache $true -CleanOldProfiles $true -ProfileMaxAgeDays 365 -SilentShow more lines

-MinimumAgeHours 48 → safer temp cleanup window
-IncludeLogs → also purge .log/.evtx/.dmp in temp locations (optional)
-CcmCacheMaxAgeDays 32 → keeps your 32‑day policy for ccmcache
-ProfileMaxAgeDays 365 → deletes profiles unused for 1 year

🔎 Validation checklist
Logs present at C:\TempCleanup\Logs\...
[TEMP], [CCMCACHE], [PROFILES] METRIC lines show Candidates and Deleted counts/sizes.
For pilots, you’ll see many [PREVIEW] lines and Deleted -> 0.
#>

<#
ServerTempCleanup_SCCM.ps1
-----------------------------------------------------------------------
• Logs: C:\TempCleanup\Logs\TempCleanup_<COMPUTER>_<YYYYMMDD>.log
• Temp cleanup (excludes C:\Temp), age-based, optional excludes.
• SCCM Client Cache: delete items older than N days (default 32), keep root.
  - Auto-detect cache path via WMI/registry; fallback to C:\Windows\ccmcache.
• Old user profile cleanup: remove profiles not used in ≥ 365 days.
• Supports -WhatIf for pilot runs (temp + ccmcache + profiles).
• Exit code: 0=success, 1=completed with error(s).
#>
<#
ServerTempCleanup_SCCM_10param.ps1
--------------------------------------------------------------------------------
• ≤10 parameters for SCCM Run Scripts compatibility.
• Logs: C:\TempCleanup\Logs\TempCleanup_<COMPUTER>_<YYYYMMDD>.log
• Temp cleanup (excludes C:\Temp), age-based, optional includes and -IncludeLogs.
• SCCM Client Cache: delete items older than N days (default 32), keep root.
  - Auto-override path with -CcmCachePath if needed.
• Old user profiles: remove profiles not used in ≥ 365 days.
• Supports -WhatIf for pilots; use -ListFiles to print [PREVIEW] items.
• Exit code: 0=success, 1=completed with error(s).
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
param(
    # -------- Temp cleanup options (C:\Temp is intentionally EXCLUDED) --------
    [int]$MinimumAgeHours = 0,
    [switch]$IncludeLogs,
    [string[]]$AdditionalPaths = @(),

    # -------- Logging / output --------
    [string]$LogDirectory = 'C:\TempCleanup\Logs',
    [switch]$Silent,
    [switch]$ListFiles,

    # -------- SCCM client cache (age-based) --------
    [int]$CcmCacheMaxAgeDays = 32,
    [string]$CcmCachePath,   # if empty, auto-detect

    # -------- Old profile cleanup (age-based) --------
    [int]$ProfileMaxAgeDays = 365
)

# -------------------- Setup & Metrics --------------------
$ErrorActionPreference = 'Stop'
$script:ErrorCount = 0

# Temp metrics
$script:TmpCandFiles = 0; $script:TmpCandDirs = 0; $script:TmpCandBytes = 0
$script:TmpDelFiles  = 0; $script:TmpDelDirs  = 0; $script:TmpDelBytes  = 0

# CCMCache metrics
$script:CcmCandFiles = 0; $script:CcmCandDirs = 0; $script:CcmCandBytes = 0
$script:CcmDelFiles  = 0; $script:CcmDelDirs  = 0; $script:CcmDelBytes  = 0

# Profiles metrics
$script:ProfCandidates = 0; $script:ProfDeleted = 0; $script:ProfBytesCand = 0; $script:ProfBytesDel = 0

# Defaults baked in (no params to keep under limit)
$ExcludePatterns = @('*.log','*.evtx','*.dmp')   # can be overridden via -IncludeLogs switch
$CleanCcmCache   = $true
$CleanOldProfiles= $true

# -------------------- Helpers --------------------
function New-Directory([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Get-Timestamp { (Get-Date).ToString('yyyy-MM-dd HH:mm:ss') }

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','METRIC')]
        [string]$Level = 'INFO'
    )
    try {
        New-Directory -Path $LogDirectory
        $global:LogFile = Join-Path $LogDirectory ("TempCleanup_{0}_{1}.log" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd'))
        $line = "[{0}] [{1}] {2}" -f (Get-Timestamp), $Level, $Message
        Add-Content -LiteralPath $global:LogFile -Value $line
        if (-not $Silent) { Write-Host $line }
    } catch { }
}

function Format-Bytes([long]$Bytes) {
    if ($Bytes -ge 1TB) { "{0:N2} TB" -f ($Bytes/1TB) }
    elseif ($Bytes -ge 1GB) { "{0:N2} GB" -f ($Bytes/1GB) }
    elseif ($Bytes -ge 1MB) { "{0:N2} MB" -f ($Bytes/1MB) }
    elseif ($Bytes -ge 1KB) { "{0:N2} KB" -f ($Bytes/1KB) }
    else { "$Bytes B" }
}

function Expand-Env([string]$Path) { [Environment]::ExpandEnvironmentVariables($Path) }

function Get-DirectorySize([string]$Path) {
    try {
        (Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
    } catch { 0 }
}

function Get-DefaultTempPaths {
    $paths = @()
    # OS/service temp (C:\Temp intentionally NOT included per request)
    $paths += "$env:windir\Temp"
    $paths += "C:\Windows\ServiceProfiles\LocalService\AppData\Local\Temp"
    $paths += "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Temp"

    # All user profile temp locations
    $userRoot = 'C:\Users'
    if (Test-Path $userRoot) {
        Get-ChildItem -LiteralPath $userRoot -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $paths += (Join-Path $_.FullName 'AppData\Local\Temp')
        }
    }

    $paths | ForEach-Object { Expand-Env $_ } | Select-Object -Unique
}

function Resolve-InputPaths([string[]]$Paths) {
    $resolved = New-Object System.Collections.Generic.List[string]
    foreach ($p in $Paths) {
        $expanded = Expand-Env $p
        if (Test-Path -LiteralPath $expanded) { $resolved.Add($expanded) }
    }
    $resolved | Select-Object -Unique
}

function Should-ExcludeFile($file) {
    if ($IncludeLogs) { return $false }
    foreach ($pattern in $ExcludePatterns) {
        if ($file.Name -like $pattern) { return $true }
    }
    return $false
}

function Get-CcmCachePath([string]$Override) {
    if ($Override -and $Override.Trim()) { return $Override }
    # WMI/CIM detection
    try {
        $inst = Get-CimInstance -Namespace 'root\ccm\SoftMgmtAgent' -ClassName 'CCM_CacheConfig' -ErrorAction Stop
        if ($inst -and $inst.Location) { return $inst.Location }
    } catch {}
    # Registry fallback
    try {
        $loc = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\CCM\Cache' -Name 'Location' -ErrorAction Stop).Location
        if ($loc) { return $loc }
    } catch {}
    # Default
    return 'C:\Windows\ccmcache'
}

# -------------------- TEMP CLEANUP --------------------
try {
    $tempCutoff = (Get-Date).AddHours(-1 * [Math]::Abs($MinimumAgeHours))
    Write-Log "Starting Temp Cleanup | Cutoff: $tempCutoff | WhatIf: $WhatIfPreference" 'INFO'

    $allPaths = @()
    $allPaths += Get-DefaultTempPaths          # excludes C:\Temp
    if ($AdditionalPaths -and $AdditionalPaths.Count -gt 0) { $allPaths += $AdditionalPaths }
    $allPaths = $allPaths | Select-Object -Unique

    $targetDirs = Resolve-InputPaths -Paths $allPaths
    if (-not $targetDirs -or $targetDirs.Count -eq 0) {
        Write-Log "No existing target temp directories. Skipping temp cleanup." 'WARN'
    } else {
        Write-Log ("Temp target directories ({0}): {1}" -f $targetDirs.Count, ($targetDirs -join '; ')) 'INFO'

        foreach ($dir in $targetDirs) {
            if (-not (Test-Path -LiteralPath $dir)) { continue }
            Write-Log "Scanning temp: $dir" 'INFO'

            # Files
            try {
                $files = Get-ChildItem -LiteralPath $dir -File -Recurse -Force -ErrorAction SilentlyContinue |
                    Where-Object { $_.LastWriteTime -lt $tempCutoff -and -not (Should-ExcludeFile $_) }
            } catch {
                Write-Log "Error enumerating files in $dir : $($_.Exception.Message)" 'ERROR'
                $script:ErrorCount++
                continue
            }

            foreach ($f in $files) {
                $script:TmpCandFiles++
                $script:TmpCandBytes += $f.Length

                if ($ListFiles -and $WhatIfPreference) {
                    Write-Log ("[PREVIEW] File: {0} ({1})" -f $f.FullName, (Format-Bytes $f.Length)) 'METRIC'
                }

                if ($PSCmdlet.ShouldProcess($f.FullName, 'Remove file')) {
                    try {
                        $size = $f.Length
                        Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                        $script:TmpDelFiles++
                        $script:TmpDelBytes += $size
                    } catch {
                        Write-Log "Failed to remove file: $($f.FullName) : $($_.Exception.Message)" 'WARN'
                    }
                }
            }

            # Empty directories older than cutoff
            try {
                $dirs = Get-ChildItem -LiteralPath $dir -Directory -Recurse -Force -ErrorAction SilentlyContinue
            } catch {
                Write-Log "Error enumerating directories in $dir : $($_.Exception.Message)" 'ERROR'
                $script:ErrorCount++
                continue
            }

            foreach ($d in ($dirs | Sort-Object FullName -Descending)) {
                $oldEnough = $d.LastWriteTime -lt $tempCutoff
                $isEmpty = $false
                try {
                    $isEmpty = -not (Get-ChildItem -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
                } catch { continue }

                if ($oldEnough -and $isEmpty) {
                    $script:TmpCandDirs++
                    if ($ListFiles -and $WhatIfPreference) {
                        Write-Log ("[PREVIEW] EmptyDir: {0}" -f $d.FullName) 'METRIC'
                    }
                    if ($PSCmdlet.ShouldProcess($d.FullName, 'Remove empty directory')) {
                        try {
                            Remove-Item -LiteralPath $d.FullName -Force -ErrorAction Stop
                            $script:TmpDelDirs++
                        } catch {
                            Write-Log "Failed to remove directory: $($d.FullName) : $($_.Exception.Message)" 'WARN'
                        }
                    }
                }
            }
        }

        Write-Log ("[TEMP] Candidates -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:TmpCandFiles, $script:TmpCandDirs, (Format-Bytes $script:TmpCandBytes)) 'METRIC'
        Write-Log ("[TEMP] Deleted    -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:TmpDelFiles,  $script:TmpDelDirs,  (Format-Bytes $script:TmpDelBytes)) 'METRIC'
    }
}
catch {
    Write-Log "Fatal error during temp cleanup: $($_.Exception.Message)" 'ERROR'
    $script:ErrorCount++
}

# -------------------- SCCM CLIENT CACHE (age-based) --------------------
if ($true) {
    try {
        $resolvedCcmPath = Get-CcmCachePath -Override $CcmCachePath
        if (Test-Path -LiteralPath $resolvedCcmPath) {
            $ccmCutoff = (Get-Date).AddDays(-1 * [Math]::Abs($CcmCacheMaxAgeDays))
            Write-Log "Cleaning SCCM Client Cache (age-based) at: $resolvedCcmPath | Cutoff: $ccmCutoff | Keep root" 'INFO'

            # Candidates
            $ccmFiles = Get-ChildItem -LiteralPath $resolvedCcmPath -Recurse -Force -File -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -lt $ccmCutoff }
            $ccmDirsAll = Get-ChildItem -LiteralPath $resolvedCcmPath -Recurse -Force -Directory -ErrorAction SilentlyContinue

            $script:CcmCandFiles = $ccmFiles.Count
            $script:CcmCandBytes = ($ccmFiles | Measure-Object -Property Length -Sum).Sum
            $script:CcmCandDirs  = ($ccmDirsAll | Where-Object { $_.LastWriteTime -lt $ccmCutoff }).Count

            # Delete aged files
            foreach ($f in $ccmFiles) {
                if ($ListFiles -and $WhatIfPreference) {
                    Write-Log ("[PREVIEW][CCM] File: {0} ({1})" -f $f.FullName, (Format-Bytes $f.Length)) 'METRIC'
                }
                if ($PSCmdlet.ShouldProcess($f.FullName, 'Remove SCCM cache file')) {
                    try {
                        Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                        $script:CcmDelFiles++
                        $script:CcmDelBytes += $f.Length
                    } catch {
                        Write-Log "SCCM cache file failed to delete: $($f.FullName) : $($_.Exception.Message)" 'WARN'
                    }
                }
            }

            # Delete empty directories older than cutoff (bottom-up)
            foreach ($d in ($ccmDirsAll | Sort-Object FullName -Descending)) {
                if ($d.LastWriteTime -lt $ccmCutoff) {
                    $isEmpty = $false
                    try {
                        $isEmpty = -not (Get-ChildItem -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
                    } catch { $isEmpty = $false }
                    if ($isEmpty) {
                        if ($ListFiles -and $WhatIfPreference) {
                            Write-Log ("[PREVIEW][CCM] EmptyDir: {0}" -f $d.FullName) 'METRIC'
                        }
                        if ($PSCmdlet.ShouldProcess($d.FullName, 'Remove SCCM cache empty directory')) {
                            try {
                                $script:CcmDelBytes += Get-DirectorySize -Path $d.FullName
                                Remove-Item -LiteralPath $d.FullName -Force -ErrorAction Stop
                                $script:CcmDelDirs++
                            } catch {
                                Write-Log "SCCM cache directory failed to delete: $($d.FullName) : $($_.Exception.Message)" 'WARN'
                            }
                        }
                    }
                }
            }

            Write-Log ("[CCMCACHE] Candidates -> Files: {0}, Dirs(aged): {1}, Size: {2}" -f $script:CcmCandFiles, $script:CcmCandDirs, (Format-Bytes $script:CcmCandBytes)) 'METRIC'
            Write-Log ("[CCMCACHE] Deleted    -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:CcmDelFiles,  $script:CcmDelDirs,  (Format-Bytes $script:CcmDelBytes)) 'METRIC'
        } else {
            Write-Log "SCCM Client Cache path not found: $resolvedCcmPath (skipping)" 'WARN'
        }
    } catch {
        Write-Log "Fatal error during SCCM cache cleanup: $($_.Exception.Message)" 'ERROR'
        $script:ErrorCount++
    }
}

# -------------------- OLD USER PROFILES (≥ N days) --------------------
if ($true) {
    try {
        $profileCutoff = (Get-Date).AddDays(-1 * [Math]::Abs($ProfileMaxAgeDays))
        Write-Log "Removing local user profiles unused since on/before: $profileCutoff" 'INFO'

        $currentSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value

        $profiles = Get-CimInstance -ClassName Win32_UserProfile -ErrorAction SilentlyContinue |
            Where-Object {
                $_.LocalPath -like 'C:\Users\*' -and
                -not $_.Special -and
                -not $_.Loaded
            }

        foreach ($p in $profiles) {
            if ($p.SID -in @('S-1-5-18','S-1-5-19','S-1-5-20', $currentSid)) { continue }
            if ($p.LocalPath -match '\\Users\\(Default|Default User|Public)($|\\)') { continue }

            $lastUse = $null
            if ($p.LastUseTime) {
                try { $lastUse = [System.Management.ManagementDateTimeConverter]::ToDateTime($p.LastUseTime) } catch { $lastUse = $null }
            }
            if (-not $lastUse -and (Test-Path -LiteralPath $p.LocalPath)) {
                try { $lastUse = (Get-Item -LiteralPath $p.LocalPath).LastWriteTime } catch { $lastUse = $null }
            }
            if (-not $lastUse) { $lastUse = (Get-Date).AddYears(-20) }

            if ($lastUse -le $profileCutoff) {
                $script:ProfCandidates++
                $size = 0
                if (Test-Path -LiteralPath $p.LocalPath) { $size = Get-DirectorySize -Path $p.LocalPath }
                $script:ProfBytesCand += $size

                if ($ListFiles -and $WhatIfPreference) {
                    Write-Log ("[PREVIEW][PROFILE] SID={0} Path={1} LastUse={2} Size≈{3}" -f $p.SID, $p.LocalPath, $lastUse, (Format-Bytes $size)) 'METRIC'
                }

                if ($PSCmdlet.ShouldProcess($p.LocalPath, "Delete user profile SID=$($p.SID)")) {
                    try {
                        $result = Invoke-CimMethod -InputObject $p -MethodName Delete -ErrorAction Stop
                        if ($result.ReturnValue -eq 0) {
                            $script:ProfDeleted++
                            $script:ProfBytesDel += $size
                            Write-Log ("Deleted user profile: SID={0} Path={1} LastUse={2}" -f $p.SID, $p.LocalPath, $lastUse) 'INFO'
                        } else {
                            Write-Log ("Failed to delete profile (code {0}): SID={1} Path={2}" -f $result.ReturnValue, $p.SID, $p.LocalPath) 'WARN'
                        }
                    } catch {
                        Write-Log ("Error deleting profile: SID={0} Path={1} : {2}" -f $p.SID, $p.LocalPath, $_.Exception.Message) 'WARN'
                    }
                }
            }
        }

        Write-Log ("[PROFILES] Candidates -> Count: {0}, Size: {1}" -f $script:ProfCandidates, (Format-Bytes $script:ProfBytesCand)) 'METRIC'
        Write-Log ("[PROFILES] Deleted    -> Count: {0}, Size: {1}" -f $script:ProfDeleted,   (Format-Bytes $script:ProfBytesDel))  'METRIC'
    } catch {
        Write-Log "Fatal error during profile cleanup: $($_.Exception.Message)" 'ERROR'
        $script:ErrorCount++
    }
}

# -------------------- FINAL --------------------
if ($WhatIfPreference) {
    Write-Log "Ran in WhatIf (preview) mode. No deletions were performed." 'INFO'
}

if ($script:ErrorCount -gt 0) {
    Write-Log "Completed with $script:ErrorCount error(s)." 'WARN'
    exit 1
} else {
    Write-Log "Completed successfully." 'INFO'
    exit 0
}
``

