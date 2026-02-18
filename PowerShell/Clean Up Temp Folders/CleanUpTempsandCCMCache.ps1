param(
    [string[]]$AdditionalPaths,
    [int]$DeleteOlderThanDays = 7,
    [switch]$ClearCCMCache
)

$LogPath = "C:\TempCleanup\Logs\TempCleanup_$(Get-Date -Format yyyyMMdd_HHmmss).log"

$DefaultPaths = @(
    "C:\Temp",
    "C:\Windows\Temp",
    "C:\Users\*\AppData\Local\Temp"
)

$PathsToClean = $DefaultPaths
if ($AdditionalPaths) {
    $PathsToClean += $AdditionalPaths
}

function Write-Log {
    param ($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$TimeStamp - $Message" | Out-File -FilePath $LogPath -Append -Encoding utf8
}

Write-Log "===== Server Cleanup Started ====="
Write-Log "DeleteOlderThanDays: $DeleteOlderThanDays"
Write-Log "ClearCCMCache: $ClearCCMCache"

# =========================
# TEMP CLEANUP
# =========================

foreach ($Path in $PathsToClean) {

    if (-not (Test-Path $Path)) {
        Write-Log "Path not found: $Path"
        continue
    }

    Write-Log "Processing: $Path"

    try {
        $Cutoff = (Get-Date).AddDays(-$DeleteOlderThanDays)

        $Items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
                 Where-Object { $_.LastWriteTime -lt $Cutoff }

        foreach ($Item in $Items | Sort-Object FullName -Descending) {
            try {
                Remove-Item $Item.FullName -Recurse -Force -ErrorAction Stop
                Write-Log "Removed: $($Item.FullName)"
            }
            catch {
                Write-Log "Skipped: $($Item.FullName)"
            }
        }
    }
    catch {
        Write-Log "Error processing ${Path}: $_"
    }
}

# =========================
# FULL CCMCACHE WIPE (RELIABLE METHOD)
# =========================

if ($ClearCCMCache) {

    Write-Log "Starting FULL CCMCache cleanup..."

    try {

        Write-Log "Stopping CcmExec service..."
        Stop-Service CcmExec -Force -ErrorAction Stop
        Start-Sleep -Seconds 5

        $CachePath = "C:\Windows\ccmcache"

        if (Test-Path $CachePath) {
            Remove-Item $CachePath -Recurse -Force -ErrorAction Stop
            Write-Log "CCMCache folder removed."
        }

        Write-Log "Starting CcmExec service..."
        Start-Service CcmExec -ErrorAction Stop

        Write-Log "CCMCache cleanup completed successfully."
    }
    catch {
        Write-Log "CCMCache cleanup failed: $_"
    }
}

Write-Log "===== Server Cleanup Completed ====="
exit 0
