# -DeleteOlderThanDays 14 -ClearCCMCache

# -ClearCCMCache
# -ClearUpdateCache
# - To remove only update cache older than 30 days:
# - DeleteOlderThanDays 30 -ClearUpdateCache


<#⚠️ What Will Happen After This Runs

CCMCache fully wiped

Required apps will re-download if needed

Temporary spike in DP traffic

C:\ space immediately reclaimed

Client recreates cache folder automatically

🏆 Enterprise Best Practice For Your Scenario

For general member servers:

Use 7–14 day temp cleanup

Full CCMCache wipe is fine

Run outside patch window

Deploy to small collection first

Monitor ccmexec.log and CacheManager.log if needed 
#>
# $LogPath = "C:\TempCleanup\Logs\TempCleanup_$(Get-Date -Format yyyyMMdd_HHmmss).log"
param(
    [string[]]$AdditionalPaths,
    [int]$DeleteOlderThanDays = 7,
    [switch]$ClearUpdateCache
)

$LogPath = "C:\TempCleanup\Logs\TempCleanup_$(Get-Date -Format yyyyMMdd_HHmmss).log"

# Default temp locations
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
Write-Log "ClearUpdateCache: $ClearUpdateCache"

# ==========================================================
# TEMP FILE CLEANUP
# ==========================================================

foreach ($Path in $PathsToClean) {

    if (-not (Test-Path $Path)) {
        Write-Log "Path not found: $Path"
        continue
    }

    Write-Log "Processing Temp Path: $Path"

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
                Write-Log "Skipped (In Use/Protected): $($Item.FullName)"
            }
        }
    }
    catch {
        Write-Log "Error processing ${Path}: $_"
    }
}

# ==========================================================
# SOFTWARE UPDATE CCMCACHE CLEANUP ONLY
# ==========================================================

if ($ClearUpdateCache) {

    Write-Log "Starting Software Update CCMCache Cleanup..."

    try {
        $UIResourceMgr = New-Object -ComObject UIResource.UIResourceMgr
        $Cache = $UIResourceMgr.GetCacheInfo()
        $CacheElements = $Cache.GetCacheElements()

        foreach ($Element in $CacheElements) {

            # ContentType 2 = Software Updates
            if ($Element.ContentType -eq 2) {

                $RemoveItem = $true

                if ($DeleteOlderThanDays -gt 0) {
                    $Cutoff = (Get-Date).AddDays(-$DeleteOlderThanDays)
                    if ($Element.LastReferenceTime -gt $Cutoff) {
                        $RemoveItem = $false
                    }
                }

                if ($RemoveItem) {
                    try {
                        $Cache.DeleteCacheElement($Element.CacheElementID)
                        Write-Log "Removed Update Cache ID: $($Element.CacheElementID)"
                    }
                    catch {
                        Write-Log "Failed to remove ID: $($Element.CacheElementID)"
                    }
                }
            }
        }

        Write-Log "Software Update CCMCache Cleanup Completed."
    }
    catch {
        Write-Log "Update Cache Cleanup Failed: $_"
    }
}

Write-Log "===== Server Cleanup Completed ====="
exit 0


