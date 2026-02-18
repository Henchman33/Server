<# 
🎯 What This Version Does Now!
For: C:\Temp
It will:
Item	Result
C:\Temp\file.tmp	✅ Deleted
C:\Temp\Junk\file.log	✅ Deleted
C:\Temp\Junk	✅ Deleted
In-use files	❌ Skipped
System-protected files	❌ Skipped
🚨 VERY IMPORTANT (Server Safety Warning)

On production servers, deleting entire folder structures can break:
App installer temp staging
Backup software temp folders
AV quarantine staging
Application cache folders
Recommended for production:
Use: -DeleteOlderThanDays 7

Example:
powershell.exe -ExecutionPolicy Bypass -File Cleanup.ps1 -DeleteOlderThanDays 7


This prevents deletion of active working folders.
#>

param(
    [switch]$WhatIfMode,
    [string[]]$AdditionalPaths,
    [int]$DeleteOlderThanDays = 0   # Set to 7+ in production if desired
)

$LogPath = "C:\TempCleanup\Logs\TempCleanup_$(Get-Date -Format yyyyMMdd_HHmmss).log"

$DefaultPaths = @(
    "C:\Temp",
    "C:\Windows\Temp",
    "C:\Users\*\AppData\Local\Temp"
)

if ($AdditionalPaths) {
    $PathsToClean = $DefaultPaths + $AdditionalPaths
}
else {
    $PathsToClean = $DefaultPaths
}

function Write-Log {
    param ($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$TimeStamp - $Message" | Out-File -FilePath $LogPath -Append -Encoding utf8
}

Write-Log "===== Temp Cleanup Started ====="
Write-Log "WhatIf Mode: $WhatIfMode"
Write-Log "DeleteOlderThanDays: $DeleteOlderThanDays"

foreach ($Path in $PathsToClean) {

    if (-not (Test-Path $Path)) {
        Write-Log "Path not found: $Path"
        continue
    }

    Write-Log "Processing: $Path"

    try {

        $Items = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue

        if ($DeleteOlderThanDays -gt 0) {
            $Cutoff = (Get-Date).AddDays(-$DeleteOlderThanDays)
            $Items = $Items | Where-Object { $_.LastWriteTime -lt $Cutoff }
        }

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

Write-Log "===== Temp Cleanup Completed ====="
exit 0
