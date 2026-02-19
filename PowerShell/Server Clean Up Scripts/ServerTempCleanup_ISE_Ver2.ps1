<# ISE‑ready script so that the SCCM client cache (C:\Windows\ccmcache) is cleaned by age:

✅ Only delete files and folders in ccmcache that are older than 32 days (configurable).
✅ Preserve the ccmcache root folder.
✅ Do NOT clean C:\Temp (as requested).
✅ Keep all other behavior from before: temp cleanup (excluding C:\Temp), old user profile removal (≥ 365 days), and logging to C:\TempCleanup\Logs with the same format/metrics.


Run: Open PowerShell ISE as Administrator, paste the script, and press F5.
If you want a different age for ccmcache, change $CcmCacheMaxAgeDays = 32 at the top.
#>
<#
ISE-Ready Temp Cleanup (Immediate Delete) with Full Logging + CCMCache Age Policy + Old Profiles
-----------------------------------------------------------------------------------------------
• Temp cleanup (excludes C:\Temp), with age and optional excludes.
• SCCM Client Cache (C:\Windows\ccmcache): delete only items older than N days (default 32), keep root folder.
• Old user profiles: remove profiles not used in ≥ 365 days.
• Logs to C:\TempCleanup\Logs with metrics.
• Run ISE as Administrator.
#>

# ==================== CONFIGURABLE OPTIONS ====================
# Temp cleanup: delete items older than this many hours (0 = delete regardless of age)
$MinimumAgeHours = 0

# Include additional temp paths if needed
$AdditionalPaths = @(
    # Example: "D:\App\Temp","E:\Cache"
)

# Exclude these patterns by default in temp cleanup; set $IncludeLogs = $true to override
$ExcludePatterns = @('*.log','*.evtx','*.dmp')

# If $true, log-like files (*.log) and others in ExcludePatterns will NOT be excluded
$IncludeLogs = $false

# SCCM client cache cleanup (age-based)
$CleanCcmCache       = $true
$CcmCachePath        = 'C:\Windows\ccmcache'   # Root preserved
$CcmCacheMaxAgeDays  = 32                      # Only delete items older than this many days

# Old local user profile cleanup (age-based)
$ProfileMaxAgeDays = 365
$CleanOldProfiles  = $true

# Silent console output (still writes to log file)
$Silent = $false

# Log directory
$LogDirectory = 'C:\TempCleanup\Logs'
# ==================== END OPTIONS ====================

# -------------------- Setup & Helpers --------------------
$ErrorActionPreference = 'Stop'
$script:ErrorCount = 0

# Temp metrics
$script:DeletedBytes = 0
$script:DeletedFiles = 0
$script:DeletedDirs  = 0
$script:CandidateBytes = 0
$script:CandidateFiles = 0
$script:CandidateDirs  = 0

# CCMCache metrics
$script:CcmCandFiles = 0
$script:CcmCandDirs  = 0
$script:CcmCandBytes = 0
$script:CcmDelFiles  = 0
$script:CcmDelDirs   = 0
$script:CcmDelBytes  = 0

# Profile metrics
$script:ProfCandidates = 0
$script:ProfDeleted    = 0
$script:ProfBytesCand  = 0
$script:ProfBytesDel   = 0

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
    } catch {
        # swallow logging errors to avoid breaking cleanup
    }
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
        $sum = Get-ChildItem -LiteralPath $Path -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
        return $sum.Sum
    } catch { return 0 }
}

function Get-DefaultTempPaths {
    $paths = @()

    # OS-level temp locations (C:\Temp intentionally NOT included per request)
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

    # Return unique, expanded
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

# -------------------- Temp Cleanup (general) --------------------
try {
    $cutoff = (Get-Date).AddHours(-1 * [Math]::Abs($MinimumAgeHours))
    Write-Log "Starting Temp Cleanup on $env:COMPUTERNAME | Cutoff: $($cutoff) | Mode: Immediate Delete" 'INFO'

    $allPaths = @()
    $allPaths += Get-DefaultTempPaths   # C:\Temp EXCLUDED by design
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
                    Where-Object { $_.LastWriteTime -lt $cutoff -and -not (Should-ExcludeFile $_) }
            } catch {
                Write-Log "Error enumerating files in $dir : $($_.Exception.Message)" 'ERROR'
                $script:ErrorCount++
                continue
            }

            foreach ($f in $files) {
                $script:CandidateFiles++
                $script:CandidateBytes += $f.Length
                try {
                    $size = $f.Length
                    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                    $script:DeletedFiles++
                    $script:DeletedBytes += $size
                } catch {
                    Write-Log "Failed to remove file: $($f.FullName) : $($_.Exception.Message)" 'WARN'
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
                $oldEnough = $d.LastWriteTime -lt $cutoff
                $isEmpty = $false
                try {
                    $isEmpty = -not (Get-ChildItem -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
                } catch { continue }

                if ($oldEnough -and $isEmpty) {
                    $script:CandidateDirs++
                    try {
                        Remove-Item -LiteralPath $d.FullName -Force -ErrorAction Stop
                        $script:DeletedDirs++
                    } catch {
                        Write-Log "Failed to remove directory: $($d.FullName) : $($_.Exception.Message)" 'WARN'
                    }
                }
            }
        }

        Write-Log ("[TEMP] Candidates -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:CandidateFiles, $script:CandidateDirs, (Format-Bytes $script:CandidateBytes)) 'METRIC'
        Write-Log ("[TEMP] Deleted    -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:DeletedFiles, $script:DeletedDirs, (Format-Bytes $script:DeletedBytes)) 'METRIC'
    }
}
catch {
    Write-Log "Fatal error during temp cleanup: $($_.Exception.Message)" 'ERROR'
    $script:ErrorCount++
}

# -------------------- SCCM Client Cache Cleanup (age-based) --------------------
if ($CleanCcmCache) {
    try {
        if (Test-Path -LiteralPath $CcmCachePath) {
            $ccmCutoff = (Get-Date).AddDays(-1 * [Math]::Abs($CcmCacheMaxAgeDays))
            Write-Log "Cleaning SCCM Client Cache (age-based): $CcmCachePath | Cutoff: $ccmCutoff | Keep root" 'INFO'

            # Measure candidates by age
            $ccmFileCandidates = Get-ChildItem -LiteralPath $CcmCachePath -Recurse -Force -File -ErrorAction SilentlyContinue |
                                 Where-Object { $_.LastWriteTime -lt $ccmCutoff }

            # For directories: we will remove only directories that are (1) older than cutoff, AND (2) empty after file deletions
            $ccmDirAll = Get-ChildItem -LiteralPath $CcmCachePath -Recurse -Force -Directory -ErrorAction SilentlyContinue

            $script:CcmCandFiles = $ccmFileCandidates.Count
            $script:CcmCandBytes = ($ccmFileCandidates | Measure-Object -Property Length -Sum).Sum
            $script:CcmCandDirs  = ($ccmDirAll | Where-Object { $_.LastWriteTime -lt $ccmCutoff }).Count  # indicative only

            # Delete files older than cutoff
            foreach ($f in $ccmFileCandidates) {
                try {
                    Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
                    $script:CcmDelFiles++
                    $script:CcmDelBytes += $f.Length
                } catch {
                    Write-Log "SCCM cache file failed to delete: $($f.FullName) : $($_.Exception.Message)" 'WARN'
                }
            }

            # Now delete empty directories that are older than cutoff (bottom-up)
            foreach ($d in ($ccmDirAll | Sort-Object FullName -Descending)) {
                if ($d.LastWriteTime -lt $ccmCutoff) {
                    $isEmpty = $false
                    try {
                        $isEmpty = -not (Get-ChildItem -LiteralPath $d.FullName -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
                    } catch { $isEmpty = $false }
                    if ($isEmpty) {
                        try {
                            # Measure size pre-delete (should be zero, but just in case)
                            $script:CcmDelBytes += Get-DirectorySize -Path $d.FullName
                            Remove-Item -LiteralPath $d.FullName -Force -ErrorAction Stop
                            $script:CcmDelDirs++
                        } catch {
                            Write-Log "SCCM cache directory failed to delete: $($d.FullName) : $($_.Exception.Message)" 'WARN'
                        }
                    }
                }
            }

            Write-Log ("[CCMCACHE] Candidates -> Files: {0}, Dirs(aged): {1}, Size: {2}" -f $script:CcmCandFiles, $script:CcmCandDirs, (Format-Bytes $script:CcmCandBytes)) 'METRIC'
            Write-Log ("[CCMCACHE] Deleted    -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:CcmDelFiles, $script:CcmDelDirs, (Format-Bytes $script:CcmDelBytes)) 'METRIC'
        } else {
            Write-Log "SCCM Client Cache path not found: $CcmCachePath (skipping)" 'WARN'
        }
    } catch {
        Write-Log "Fatal error during SCCM cache cleanup: $($_.Exception.Message)" 'ERROR'
        $script:ErrorCount++
    }
}

# -------------------- Old User Profile Cleanup --------------------
if ($CleanOldProfiles) {
    try {
        $cutoffDate = (Get-Date).AddDays(-1 * [Math]::Abs($ProfileMaxAgeDays))
        Write-Log "Removing local user profiles unused since on/before: $($cutoffDate)" 'INFO'

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

            if ($lastUse -le $cutoffDate) {
                $script:ProfCandidates++
                $size = 0
                if (Test-Path -LiteralPath $p.LocalPath) { $size = Get-DirectorySize -Path $p.LocalPath }
                $script:ProfBytesCand += $size

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

        Write-Log ("[PROFILES] Candidates -> Count: {0}, Size: {1}" -f $script:ProfCandidates, (Format-Bytes $script:ProfBytesCand)) 'METRIC'
        Write-Log ("[PROFILES] Deleted    -> Count: {0}, Size: {1}" -f $script:ProfDeleted, (Format-Bytes $script:ProfBytesDel)) 'METRIC'
    } catch {
        Write-Log "Fatal error during profile cleanup: $($_.Exception.Message)" 'ERROR'
        $script:ErrorCount++
    }
}

# -------------------- Final Summary --------------------
if ($script:ErrorCount -gt 0) {
    Write-Log "Completed with $script:ErrorCount error(s)." 'WARN'
} else {
    Write-Log "Completed successfully." 'INFO'
}
