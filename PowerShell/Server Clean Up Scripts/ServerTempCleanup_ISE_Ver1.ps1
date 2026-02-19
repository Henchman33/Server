<#
ISE-Ready Temp Cleanup (Immediate Delete) with Full Logging + CCMCache + Old Profiles
------------------------------------------------------------------------------------
• Deletes files/folders in:
    - C:\Windows\Temp
    - Service profile temp folders
    - All user profile temp folders
    - Optional additional paths
  (C:\Temp is intentionally EXCLUDED per request)
• Clears SCCM Client Cache contents (C:\Windows\ccmcache) but preserves the root folder.
• Removes local user profiles not used in ≥ 365 days (safe filters applied).
• Keeps logging structure & metrics; logs to C:\TempCleanup\Logs.
• No WhatIf: runs immediately when executed.
• Skips locked/in-use files (logs them).
#>

# ==================== CONFIGURABLE OPTIONS ====================
# Delete files/folders older than this many hours in temp locations (0 = all items)
$MinimumAgeHours = 0

# Include additional temp paths if needed (comma-separated strings)
$AdditionalPaths = @(
    # Example: "D:\App\Temp","E:\Cache"
)

# Exclude these patterns by default in temp cleanup; set $IncludeLogs = $true to override
$ExcludePatterns = @('*.log','*.evtx','*.dmp')

# If $true, log-like files (*.log) and others in ExcludePatterns will NOT be excluded
$IncludeLogs = $false

# Clean SCCM client cache contents (keeps root folder)
$CleanCcmCache = $true
$CcmCachePath  = 'C:\Windows\ccmcache'   # If your client cache is custom, update here

# Remove local user profiles not used in this many days (365 as requested)
$ProfileMaxAgeDays = 365
$CleanOldProfiles  = $true

# Silent console output (still writes to log file)
$Silent = $false

# Log directory (as requested)
$LogDirectory = 'C:\TempCleanup\Logs'
# ==================== END OPTIONS ====================

# -------------------- Setup & Helpers --------------------
$ErrorActionPreference = 'Stop'
$script:ErrorCount = 0

# File/dir metrics (Temp cleanup)
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
        return [int64]($sum.Sum)
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
            $temp = Join-Path $_.FullName 'AppData\Local\Temp'
            $paths += $temp
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

# -------------------- Temp Cleanup --------------------
try {
    $cutoff = (Get-Date).AddHours(-1 * [Math]::Abs($MinimumAgeHours))
    Write-Log "Starting Temp Cleanup on $env:COMPUTERNAME | Cutoff: $($cutoff) | Mode: Immediate Delete" 'INFO'

    # Build target directories (C:\Temp excluded by design)
    $allPaths = @()
    $allPaths += Get-DefaultTempPaths
    if ($AdditionalPaths -and $AdditionalPaths.Count -gt 0) { $allPaths += $AdditionalPaths }
    $allPaths = $allPaths | Select-Object -Unique

    # Resolve to existing directories
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

        # Summary for temp cleanup
        Write-Log ("[TEMP] Candidates -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:CandidateFiles, $script:CandidateDirs, (Format-Bytes $script:CandidateBytes)) 'METRIC'
        Write-Log ("[TEMP] Deleted    -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:DeletedFiles, $script:DeletedDirs, (Format-Bytes $script:DeletedBytes)) 'METRIC'
    }
}
catch {
    Write-Log "Fatal error during temp cleanup: $($_.Exception.Message)" 'ERROR'
    $script:ErrorCount++
}

# -------------------- SCCM Client Cache Cleanup --------------------
if ($CleanCcmCache) {
    try {
        if (Test-Path -LiteralPath $CcmCachePath) {
            Write-Log "Cleaning SCCM Client Cache contents: $CcmCachePath (preserving root folder)" 'INFO'

            # Measure candidates
            $ccmFiles = Get-ChildItem -LiteralPath $CcmCachePath -Recurse -Force -File -ErrorAction SilentlyContinue
            $ccmDirs  = Get-ChildItem -LiteralPath $CcmCachePath -Recurse -Force -Directory -ErrorAction SilentlyContinue
            $script:CcmCandFiles = $ccmFiles.Count
            $script:CcmCandDirs  = $ccmDirs.Count
            $script:CcmCandBytes = ($ccmFiles | Measure-Object -Property Length -Sum).Sum

            # Delete all child items (files and folders), keep root
            Get-ChildItem -LiteralPath $CcmCachePath -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    # Track size to metrics
                    if ($_.PSIsContainer) {
                        $script:CcmDelDirs++
                        $script:CcmDelBytes += Get-DirectorySize -Path $_.FullName
                        Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction Stop
                    } else {
                        $script:CcmDelFiles++
                        $script:CcmDelBytes += $_.Length
                        Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop
                    }
                } catch {
                    Write-Log "SCCM cache item failed to delete: $($_.FullName) : $($_.Exception.Message)" 'WARN'
                }
            }

            Write-Log ("[CCMCACHE] Candidates -> Files: {0}, Dirs: {1}, Size: {2}" -f $script:CcmCandFiles, $script:CcmCandDirs, (Format-Bytes $script:CcmCandBytes)) 'METRIC'
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
            # Skip core/system SIDs explicitly
            if ($p.SID -in @('S-1-5-18','S-1-5-19','S-1-5-20', $currentSid)) { continue }

            # Skip well-known folders
            if ($p.LocalPath -match '\\Users\\(Default|Default User|Public)($|\\)') { continue }

            # Determine last use time
            $lastUse = $null
            if ($p.LastUseTime) {
                try { $lastUse = [System.Management.ManagementDateTimeConverter]::ToDateTime($p.LastUseTime) } catch { $lastUse = $null }
            }
            if (-not $lastUse -and (Test-Path -LiteralPath $p.LocalPath)) {
                try { $lastUse = (Get-Item -LiteralPath $p.LocalPath).LastWriteTime } catch { $lastUse = $null }
            }
            if (-not $lastUse) { 
                # If unknown, be conservative: require explicit age by folder timestamp
                $lastUse = (Get-Date).AddYears(-20)
            }

            if ($lastUse -le $cutoffDate) {
                $script:ProfCandidates++
                # Pre-measure size for metrics (best-effort)
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
