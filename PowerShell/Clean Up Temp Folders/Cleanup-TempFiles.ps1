#Cleanup-TempFiles.ps1<#
.SYNOPSIS
    Silent temp file cleanup script for servers (ConfigMgr safe)

.DESCRIPTION
    Cleans common temp locations and allows custom paths.
    Supports -WhatIf mode for safe testing.
    Logs results locally.

.PARAMETER WhatIfMode
    Runs in test mode without deleting files.

.PARAMETER AdditionalPaths
    Array of additional folders to clean.

.EXAMPLE
    .\Cleanup-TempFiles.ps1 -WhatIfMode

.EXAMPLE
    .\Cleanup-TempFiles.ps1 -AdditionalPaths "D:\Temp","E:\AppCache"
#>

param(
    [switch]$WhatIfMode,
    [string[]]$AdditionalPaths
)

# ====== Configuration ======

$LogPath = "C:\TempCleanup\Logs\TempCleanup_$(Get-Date -Format yyyyMMdd_HHmmss).log"

# Default safe temp locations
$DefaultPaths = @(
    "C:\Windows\Temp",
    #"C:\Windows\SoftwareDistribution\Download",
    "C:\Temp",
    "C:\Users\*\AppData\Local\Temp"
)

# Merge additional paths
if ($AdditionalPaths) {
    $PathsToClean = $DefaultPaths + $AdditionalPaths
}
else {
    $PathsToClean = $DefaultPaths
}

# ====== Logging Function ======

function Write-Log {
    param ($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$TimeStamp - $Message" | Out-File -FilePath $LogPath -Append -Encoding utf8
}

Write-Log "===== Temp Cleanup Script Started ====="
Write-Log "WhatIf Mode: $WhatIfMode"

# ====== Cleanup Function ======

function Remove-TempFiles {
    param ($Path)

    try {
        if (Test-Path $Path) {
            Write-Log "Processing: $Path"

            Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
                Where-Object {
                    -not $_.PSIsContainer
                } |
                ForEach-Object {
                    try {
                        if ($WhatIfMode) {
                            Write-Log "WHATIF: Would remove $($_.FullName)"
                        }
                        else {
                            Remove-Item $_.FullName -Force -ErrorAction Stop
                            Write-Log "Removed: $($_.FullName)"
                        }
                    }
                    catch {
                        Write-Log "Skipped (In Use or Protected): $($_.FullName)"
                    }
                }
        }
        else {
            Write-Log "Path not found: $Path"
        }
    }
    catch {
        Write-Log "Error processing ${Path}: $_"
    }
}

# ====== Execute Cleanup ======

foreach ($Path in $PathsToClean) {
    Remove-TempFiles -Path $Path
}

Write-Log "===== Temp Cleanup Completed ====="
exit 0
