<#
.SYNOPSIS
    Generates a comprehensive report of server drives, shares, and permissions
.DESCRIPTION
    This script collects information about local drives, shared folders, and 
    NTFS/share permissions for migration purposes
.NOTES
    Author: System Administrator
    Requires: Administrative privileges
#>

# Create output directory on desktop
$desktopPath = [Environment]::GetFolderPath("Desktop")
$outputFolder = Join-Path -Path $desktopPath -ChildPath "Shares"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Create folder if it doesn't exist
if (!(Test-Path -Path $outputFolder)) {
    New-Item -Path $outputFolder -ItemType Directory -Force | Out-Null
}

# Define output file paths
$drivesReport = Join-Path -Path $outputFolder -ChildPath "Drives_$timestamp.csv"
$sharesReport = Join-Path -Path $outputFolder -ChildPath "Shares_$timestamp.csv"
$sharePermissionsReport = Join-Path -Path $outputFolder -ChildPath "SharePermissions_$timestamp.csv"
$summaryReport = Join-Path -Path $outputFolder -ChildPath "Summary_$timestamp.csv"

Write-Host "Starting server migration report..." -ForegroundColor Green
Write-Host "Output folder: $outputFolder" -ForegroundColor Yellow

# 1. Get Drive Information
Write-Host "Collecting drive information..." -ForegroundColor Cyan
$drives = Get-WmiObject Win32_LogicalDisk | Where-Object { $_.DriveType -in @(3, 4) } | Select-Object `
    @{Name="ServerName";Expression={$env:COMPUTERNAME}},
    @{Name="DriveLetter";Expression={$_.DeviceID}},
    @{Name="VolumeName";Expression={$_.VolumeName}},
    @{Name="FileSystem";Expression={$_.FileSystem}},
    @{Name="Size(GB)";Expression={[math]::Round($_.Size / 1GB, 2)}},
    @{Name="FreeSpace(GB)";Expression={[math]::Round($_.FreeSpace / 1GB, 2)}},
    @{Name="UsedSpace(GB)";Expression={[math]::Round(($_.Size - $_.FreeSpace) / 1GB, 2)}},
    @{Name="PercentFree";Expression={[math]::Round(($_.FreeSpace / $_.Size) * 100, 2)}}

$drives | Export-Csv -Path $drivesReport -NoTypeInformation -Encoding UTF8
Write-Host "Drive information exported successfully" -ForegroundColor Green

# 2. Get Shared Folders Information
Write-Host "Collecting share information..." -ForegroundColor Cyan

# Get all non-administrative shares
$shares = Get-WmiObject Win32_Share | Where-Object { 
    $_.Name -notin @("ADMIN$", "C$", "D$", "E$", "F$", "G$", "H$", "IPC$", "print$") -and 
    $_.Type -eq 0  # Type 0 = Disk Drive shares
} | Select-Object `
    @{Name="ServerName";Expression={$env:COMPUTERNAME}},
    @{Name="ShareName";Expression={$_.Name}},
    @{Name="Path";Expression={$_.Path}},
    @{Name="Description";Expression={$_.Description}},
    @{Name="ShareType";Expression={switch($_.Type){0{"Disk Drive"}1{"Print Queue"}2{"Device"}3{"IPC"}2147483648{"Special"}}}}

$shares | Export-Csv -Path $sharesReport -NoTypeInformation -Encoding UTF8
Write-Host "Share information exported successfully" -ForegroundColor Green

# 3. Get Share Permissions for each share
Write-Host "Collecting share permissions..." -ForegroundColor Cyan

$sharePermissions = @()

foreach ($share in $shares) {
    Write-Host "Processing permissions for: $($share.ShareName)" -ForegroundColor Gray
    
    try {
        # Get share permissions using WMI
        $sharePath = $share.Path
        $shareName = $share.ShareName
        
        # Get local share permissions via net share command
        $netShareOutput = net share $shareName
        $permissions = $null
        
        # Alternative method using Get-SmbShare (PowerShell 3.0+)
        if (Get-Command Get-SmbShare -ErrorAction SilentlyContinue) {
            $smbShare = Get-SmbShare -Name $shareName -ErrorAction SilentlyContinue
            if ($smbShare) {
                $smbPermissions = Get-SmbShareAccess -Name $shareName -ErrorAction SilentlyContinue
                foreach ($perm in $smbPermissions) {
                    $sharePermissions += [PSCustomObject]@{
                        ServerName = $env:COMPUTERNAME
                        ShareName = $shareName
                        Path = $sharePath
                        UserOrGroup = $perm.AccountName
                        PermissionType = "Share Level"
                        AccessRight = $perm.AccessRight
                        IsInherited = $false
                    }
                }
            }
        }
        
        # Also get NTFS permissions for the folder (if it exists)
        if (Test-Path $sharePath) {
            $acl = Get-Acl -Path $sharePath -ErrorAction SilentlyContinue
            if ($acl) {
                foreach ($access in $acl.Access) {
                    $sharePermissions += [PSCustomObject]@{
                        ServerName = $env:COMPUTERNAME
                        ShareName = $shareName
                        Path = $sharePath
                        UserOrGroup = $access.IdentityReference
                        PermissionType = "NTFS"
                        AccessRight = $access.FileSystemRights
                        IsInherited = $access.IsInherited
                    }
                }
            }
        }
        else {
            Write-Warning "Path $sharePath for share $shareName does not exist or is inaccessible"
        }
    }
    catch {
        Write-Warning "Error processing share $($share.ShareName): $($_.Exception.Message)"
    }
}

# Export share permissions
if ($sharePermissions.Count -gt 0) {
    $sharePermissions | Export-Csv -Path $sharePermissionsReport -NoTypeInformation -Encoding UTF8
    Write-Host "Share permissions exported successfully" -ForegroundColor Green
} else {
    "No share permissions found" | Out-File -FilePath $sharePermissionsReport
}

# 4. Create Summary Report
Write-Host "Creating summary report..." -ForegroundColor Cyan

$summary = [PSCustomObject]@{
    ServerName = $env:COMPUTERNAME
    ReportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    TotalDrives = $drives.Count
    TotalShares = $shares.Count
    TotalSharePermissions = $sharePermissions.Count
    DrivesReportFile = $drivesReport
    SharesReportFile = $sharesReport
    PermissionsReportFile = $sharePermissionsReport
}

$summary | Export-Csv -Path $summaryReport -NoTypeInformation -Encoding UTF8

# 5. Display Summary
Write-Host "`n" + "="*60 -ForegroundColor Green
Write-Host "REPORT GENERATION COMPLETE!" -ForegroundColor Green
Write-Host "="*60 -ForegroundColor Green
Write-Host "Server: $($summary.ServerName)" -ForegroundColor Yellow
Write-Host "Report Date: $($summary.ReportDate)" -ForegroundColor Yellow
Write-Host "Total Drives: $($summary.TotalDrives)" -ForegroundColor Yellow
Write-Host "Total Shares: $($summary.TotalShares)" -ForegroundColor Yellow
Write-Host "Total Permissions: $($summary.TotalSharePermissions)" -ForegroundColor Yellow
Write-Host "`nFiles created:" -ForegroundColor Cyan
Write-Host "1. $($drivesReport)" -ForegroundColor White
Write-Host "2. $($sharesReport)" -ForegroundColor White
Write-Host "3. $($sharePermissionsReport)" -ForegroundColor White
Write-Host "4. $($summaryReport)" -ForegroundColor White
Write-Host "`nAll reports saved to: $outputFolder" -ForegroundColor Green
Write-Host "="*60 -ForegroundColor Green

# Optional: Open the folder in Explorer
Start-Process explorer.exe $outputFolder

Write-Host "`nPress any key to exit..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
