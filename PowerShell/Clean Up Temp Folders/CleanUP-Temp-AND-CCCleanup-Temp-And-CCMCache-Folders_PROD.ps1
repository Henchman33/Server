# Cleanup-Temp-And-CCMCache.ps1
<#🚀 How To Deploy (Recommended Method)
Option 1 – Run Script (Fastest)
Use parameter:
-ClearCCMCache
Optional:
-DeleteOlderThanDays 14 -ClearCCMCache
Option 2 – Package Deployment
Program:
powershell.exe -ExecutionPolicy Bypass -File Cleanup-Temp-And-CCMCache.ps1 -ClearCCMCache
Run:
Whether user is logged on or not
Hidden
With administrative rights
⚠️ Production Considerations
Clearing CCMCache will:
Remove all cached deployment content
Force re-download of required applications
Increase network traffic temporarily
I recommend: Running outside patch maintenance window
* Testing on small server collection first
* Using at least -DeleteOlderThanDays 7 #>
param(
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
                $Cache.DeleteCacheElement($Element.CacheElementID)
                Write-Log "Removed CCMCache Item ID: $($Element.CacheElementID)"
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
