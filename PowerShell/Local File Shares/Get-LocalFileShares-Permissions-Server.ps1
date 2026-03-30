#Requires -RunAsAdministrator
<#
.SYNOPSIS
    Generates a comprehensive report of all drives, shared folders, and share permissions
    on a local Windows Server within an Active Directory domain.

.DESCRIPTION
    This script collects:
      - All local fixed drives (size, free space, file system)
      - All SMB shared folders (name, path, description, type)
      - Share permissions (ACL) for each share — trustees, access rights, access type
    Results are exported as two CSV files to the running user's Desktop under a
    folder named "Shares":
      1. DriveReport_<ServerName>_<Date>.csv
      2. SharePermissions_<ServerName>_<Date>.csv

.NOTES
    - Must be run as a Domain/Local Administrator.
    - Requires the SmbShare module (available on Windows Server 2012 R2+).
    - Run from PowerShell ISE or a standard PowerShell console.

    Author : Migration Admin
    Version: 1.0
#>

# ─────────────────────────────────────────────
#  CONFIGURATION
# ─────────────────────────────────────────────
$ServerName  = $env:COMPUTERNAME
$DateStamp   = Get-Date -Format "yyyy-MM-dd_HHmm"
$DesktopPath = [Environment]::GetFolderPath("Desktop")
$OutputDir   = Join-Path $DesktopPath "Shares"

# ─────────────────────────────────────────────
#  HELPER: Ensure output directory exists
# ─────────────────────────────────────────────
if (-not (Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
    Write-Host "[INFO] Created output folder: $OutputDir" -ForegroundColor Cyan
} else {
    Write-Host "[INFO] Output folder already exists: $OutputDir" -ForegroundColor Cyan
}

$DriveCSV = Join-Path $OutputDir "DriveReport_${ServerName}_${DateStamp}.csv"
$ShareCSV = Join-Path $OutputDir "SharePermissions_${ServerName}_${DateStamp}.csv"

# ─────────────────────────────────────────────
#  SECTION 1 — DRIVE INVENTORY
# ─────────────────────────────────────────────
Write-Host "`n[STEP 1] Collecting drive information..." -ForegroundColor Yellow

$DriveReport = @()

Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -match '^[A-Z]:\\$' } | ForEach-Object {
    $Drive = $_

    # Pull WMI disk info for richer detail
    $DriveLetter = ($Drive.Root -replace '\\','')
    $WmiDisk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='$DriveLetter'" -ErrorAction SilentlyContinue

    $TotalGB = if ($WmiDisk.Size)      { [math]::Round($WmiDisk.Size       / 1GB, 2) } else { "N/A" }
    $FreeGB  = if ($WmiDisk.FreeSpace) { [math]::Round($WmiDisk.FreeSpace  / 1GB, 2) } else { "N/A" }
    $UsedGB  = if ($TotalGB -ne "N/A" -and $FreeGB -ne "N/A") { [math]::Round($TotalGB - $FreeGB, 2) } else { "N/A" }
    $PctFree = if ($TotalGB -ne "N/A" -and $TotalGB -gt 0)    { [math]::Round(($FreeGB / $TotalGB) * 100, 1) } else { "N/A" }

    $DriveType = switch ($WmiDisk.DriveType) {
        0  { "Unknown" }
        1  { "No Root Directory" }
        2  { "Removable" }
        3  { "Fixed (Local)" }
        4  { "Network" }
        5  { "Compact Disc" }
        6  { "RAM Disk" }
        default { "Unknown" }
    }

    $DriveReport += [PSCustomObject]@{
        ServerName       = $ServerName
        DriveLetter      = $DriveLetter
        DriveType        = $DriveType
        FileSystem       = $WmiDisk.FileSystem
        VolumeLabel      = $WmiDisk.VolumeName
        TotalSizeGB      = $TotalGB
        UsedSizeGB       = $UsedGB
        FreeSizeGB       = $FreeGB
        FreeSpacePct     = $PctFree
        DriveDescription = $WmiDisk.Description
    }

    Write-Host "  [+] Drive $DriveLetter  Total: ${TotalGB}GB  Free: ${FreeGB}GB ($PctFree%)" -ForegroundColor Gray
}

$DriveReport | Export-Csv -Path $DriveCSV -NoTypeInformation -Encoding UTF8
Write-Host "[OK] Drive report saved -> $DriveCSV" -ForegroundColor Green

# ─────────────────────────────────────────────
#  SECTION 2 — SHARE PERMISSIONS
# ─────────────────────────────────────────────
Write-Host "`n[STEP 2] Collecting SMB shares and permissions..." -ForegroundColor Yellow

# Retrieve all SMB shares; exclude common hidden admin shares if desired.
# Remove the Where-Object line below to include ADMIN$, C$, IPC$, etc.
$AllShares = Get-SmbShare -ErrorAction Stop | Where-Object {
    $_.Name -notmatch '^\w\$$' -and          # Excludes drive-letter admin shares (C$, D$…)
    $_.Name -notin @('ADMIN$','IPC$','print$') # Excludes other default admin shares
}

Write-Host "  Found $($AllShares.Count) non-administrative shares." -ForegroundColor Gray

$SharePermissions = @()

foreach ($Share in $AllShares) {

    Write-Host "  [+] Processing share: $($Share.Name)  Path: $($Share.Path)" -ForegroundColor Gray

    # Determine share type label
    $ShareTypeName = switch ($Share.ShareType) {
        "FileSystemDirectory" { "Folder" }
        "Printer"             { "Printer" }
        "Device"              { "Device" }
        "IPC"                 { "IPC" }
        default               { $Share.ShareType }
    }

    # ── SMB Share-level permissions ──────────────────────────────────────────
    $SmbAcls = Get-SmbShareAccess -Name $Share.Name -ErrorAction SilentlyContinue

    if ($SmbAcls) {
        foreach ($Acl in $SmbAcls) {
            $SharePermissions += [PSCustomObject]@{
                ServerName         = $ServerName
                ShareName          = $Share.Name
                SharePath          = $Share.Path
                ShareDescription   = $Share.Description
                ShareType          = $ShareTypeName
                SpecialShares      = $Share.Special
                PermissionLayer    = "Share (SMB)"
                AccountName        = $Acl.AccountName
                AccessControlType  = $Acl.AccessControlType   # Allow / Deny
                AccessRight        = $Acl.AccessRight          # Full / Change / Read
                # NTFS columns left blank for share-level rows
                NTFSInheritance    = ""
                NTFSIsInherited    = ""
            }
        }
    } else {
        # Share exists but returned no ACL entries — record the share at minimum
        $SharePermissions += [PSCustomObject]@{
            ServerName         = $ServerName
            ShareName          = $Share.Name
            SharePath          = $Share.Path
            ShareDescription   = $Share.Description
            ShareType          = $ShareTypeName
            SpecialShares      = $Share.Special
            PermissionLayer    = "Share (SMB)"
            AccountName        = "(No ACL entries returned)"
            AccessControlType  = ""
            AccessRight        = ""
            NTFSInheritance    = ""
            NTFSIsInherited    = ""
        }
    }

    # ── NTFS filesystem-level permissions (if path exists and is a folder) ──
    if ($Share.Path -and (Test-Path $Share.Path -PathType Container -ErrorAction SilentlyContinue)) {
        try {
            $Acl = Get-Acl -Path $Share.Path -ErrorAction Stop
            foreach ($Rule in $Acl.Access) {
                $SharePermissions += [PSCustomObject]@{
                    ServerName         = $ServerName
                    ShareName          = $Share.Name
                    SharePath          = $Share.Path
                    ShareDescription   = $Share.Description
                    ShareType          = $ShareTypeName
                    SpecialShares      = $Share.Special
                    PermissionLayer    = "NTFS (FileSystem)"
                    AccountName        = $Rule.IdentityReference.Value
                    AccessControlType  = $Rule.AccessControlType        # Allow / Deny
                    AccessRight        = $Rule.FileSystemRights          # e.g. FullControl, Modify…
                    NTFSInheritance    = $Rule.InheritanceFlags           # ContainerInherit / ObjectInherit
                    NTFSIsInherited    = $Rule.IsInherited                 # True / False
                }
            }
        } catch {
            Write-Warning "  Could not read NTFS ACL for '$($Share.Path)': $_"
        }
    }
}

$SharePermissions | Export-Csv -Path $ShareCSV -NoTypeInformation -Encoding UTF8
Write-Host "[OK] Share permissions report saved -> $ShareCSV" -ForegroundColor Green

# ─────────────────────────────────────────────
#  SUMMARY
# ─────────────────────────────────────────────
Write-Host "`n════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  REPORT COMPLETE — Server: $ServerName" -ForegroundColor Cyan
Write-Host "════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Drives found         : $($DriveReport.Count)"
Write-Host "  Shares processed     : $($AllShares.Count)"
Write-Host "  Permission rows      : $($SharePermissions.Count)"
Write-Host ""
Write-Host "  Output folder  : $OutputDir"
Write-Host "  Drive CSV      : $(Split-Path $DriveCSV -Leaf)"
Write-Host "  Share CSV      : $(Split-Path $ShareCSV -Leaf)"
Write-Host "════════════════════════════════════════════`n" -ForegroundColor Cyan

# Optional: open the output folder automatically
# Start-Process explorer.exe $OutputDir
