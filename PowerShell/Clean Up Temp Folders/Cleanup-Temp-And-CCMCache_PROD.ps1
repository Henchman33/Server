<#
if you're pushing this from Microsoft Configuration Manager, adding ConfigMgr client cache cleanup (ccmcache) is absolutely smart.

Below is a production-safe enterprise version that:

✅ Cleans temp folders

✅ Removes subfolders

✅ Clears ConfigMgr client cache using the supported COM method

✅ Supports -WhatIfMode

✅ Supports -DeleteOlderThanDays

✅ Logs everything

✅ Runs silently

✅ Safe for one-time SCCM push

⚠️ Important (Why This Is Safe)

We do NOT manually delete:

C:\Windows\ccmcache


Instead, we use the supported client interface:

UIResource.UIResourceMgr


This prevents:

Broken deployments

Corrupt cache state

Client repair scenarios
#>

# Cleanup-Temp-And-CCMCache.ps1

param(
    [switch]$WhatIfMode,
    [string[]]$AdditionalPaths,
    [int]$DeleteOlderThanDays = 7,
    [switch]$ClearCCMCache
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
Write-Log "WhatIf Mode: $WhatIfMode"
Write-Log "DeleteOlderThanDays: $DeleteOlderThanDays"
Write-Log "ClearCCMCache: $ClearCCMCache"

# =========================
# TEMP CLEANUP SECTION
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
                if ($WhatIfMode) {
                    Write-Log "WHATIF: Would remove $($Item.FullName)"
                }
                else {
                    Remove-Item $Item.FullName -Recurse -Force -ErrorAction Stop
                    Write-Log "Removed: $($Item.FullName)"
                }
            }
            catch {
                Write-Log "Skipped (In Use / Protected): $($Item.FullName)"
            }
        }

    }
    catch {
        Write-Log "Error processing ${Path}: $_"
    }
}

# =========================
# CONFIGMGR CLIENT CACHE CLEANUP
# =========================

if ($ClearCCMCache) {

    Write-Log "Starting CCMCache Cleanup..."

    try {
        $UIResourceMgr = New-Object -ComObject UIResource.UIResourceMgr
        $Cache = $UIResourceMgr.GetCacheInfo()
        $CacheElements = $Cache.GetCacheElements()

        foreach ($Element in $CacheElements) {

            try {
                if ($WhatIfMode) {
                    Write-Log "WHATIF: Would remove CCMCache Item ID: $($Element.CacheElementID)"
                }
                else {
                    $Cache.DeleteCacheElement($Element.CacheElementID)
                    Write-Log "Removed CCMCache Item ID: $($Element.CacheElementID)"
                }
            }
            catch {
                Write-Log "Failed to remove CCMCache Item ID: $($Element.CacheElementID)"
            }
        }

        Write-Log "CCMCache Cleanup Completed."
    }
    catch {
        Write-Log "CCMCache Cleanup Failed: $_"
    }
}

Write-Log "===== Server Cleanup Completed ====="
exit 0


