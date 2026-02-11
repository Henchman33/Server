<#
.SYNOPSIS
    Enumerates AD Windows Server computers and reports disk usage to CSV, XLSX, and HTML.

.DESCRIPTION
    - Finds server-class computers in Active Directory.
    - Collects domain (FQDN), IPv4, and disk info for each fixed drive (optionally volumes/mount points).
    - Flags low free space by GB and/or % thresholds.
    - Outputs to screen and exports CSV, XLSX, and a styled HTML report into a Desktop folder:
        "Server Low Disk Space Report" (created if missing).
    - Includes a commented-out SMTP email section (SMTP.IGT.COM).

.NOTES
    - Compatible with Windows PowerShell 5.1 on server OS.
    - Requires RSAT ActiveDirectory module.
    - Requires admin rights on servers for WMI/CIM queries.
#>

[CmdletBinding()]
param(
    [int]$WarningFreeGB = 20,
    [int]$WarningFreePct = 10,
    [switch]$ShowAll,
    [switch]$IncludeMountPoints,
    [string]$OU
)

# -------------------------
# Report folder location
# -------------------------

# Default (Option B) — Desktop of user running script:
$ReportFolder = Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath 'Server Low Disk Space Report'

# Option C (custom path) — uncomment to use and adjust location as needed
# $ReportFolder = 'D:\Reports\Server Low Disk Space Report'

# Ensure folder exists
if (-not (Test-Path -Path $ReportFolder)) {
    New-Item -Path $ReportFolder -ItemType Directory -Force | Out-Null
}

# Timestamped filenames
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$CsvPath  = Join-Path $ReportFolder "ServerDiskReport_$stamp.csv"
$XlsxPath = Join-Path $ReportFolder "ServerDiskReport_$stamp.xlsx"
$HtmlPath = Join-Path $ReportFolder "ServerDiskReport_$stamp.html"

# -------------------------
# Helper Functions
# -------------------------

function Test-Module {
    param([Parameter(Mandatory)] [string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not available."
    }
    Import-Module $Name -ErrorAction Stop | Out-Null
}

function Get-DomainFqdnFromDN {
    param([Parameter(Mandatory)] [string]$DistinguishedName)
    ($DistinguishedName -split ',') |
        Where-Object { $_ -like 'DC=*' } |
        ForEach-Object { $_.Substring(3) } |
        -join '.'
}

function Resolve-IPv4 {
    param(
        [Parameter(Mandatory)] [string]$ComputerName,
        [Parameter()] [Microsoft.Management.Infrastructure.CimSession]$CimSession
    )
    try {
        $dns = (Resolve-DnsName -Name $ComputerName -Type A -ErrorAction Stop |
            Select-Object -ExpandProperty IPAddress -First 1)
        if ($dns) { return $dns }
    } catch { }

    if ($CimSession) {
        try {
            $ip = Get-CimInstance -CimSession $CimSession -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' |
                ForEach-Object { $_.IPAddress } |
                Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' } |
                Select-Object -First 1
            if ($ip) { return $ip }
        } catch { }
    }

    return $null
}

function New-BestCimSession {
    param([Parameter(Mandatory)] [string]$ComputerName)
    try {
        return New-CimSession -ComputerName $ComputerName -ErrorAction Stop
    } catch {
        try {
            $opt = New-CimSessionOption -Protocol DCOM
            return New-CimSession -ComputerName $ComputerName -SessionOption $opt -ErrorAction Stop
        } catch {
            return $null
        }
    }
}

function Get-ServerDisks {
    param(
        [Parameter(Mandatory)] [string]$ComputerName,
        [switch]$IncludeMountPoints
    )

    $session = New-BestCimSession -ComputerName $ComputerName
    if (-not $session) {
        return [pscustomobject]@{
            ComputerName = $ComputerName
            Error        = "Unable to create CIM session"
            Disks        = @()
            CimSession   = $null
        }
    }

    try {
        if ($IncludeMountPoints) {
            $vols = Get-CimInstance -CimSession $session -ClassName Win32_Volume -Filter 'DriveType = 3' |
                Where-Object { $_.Capacity -gt 0 }

            $disks = foreach ($v in $vols) {
                $sizeGB = [math]::Round(($v.Capacity / 1GB), 2)
                $freeGB = [math]::Round(($v.FreeSpace / 1GB), 2)
                $pct    = if ($v.Capacity -gt 0) { [math]::Round(($v.FreeSpace / $v.Capacity) * 100, 2) } else { 0 }
                $name   = if ($v.DriveLetter) { $v.DriveLetter } elseif ($v.Name) { $v.Name.TrimEnd('\') } else { $v.Label }

                [pscustomobject]@{
                    Volume  = $name
                    SizeGB  = $sizeGB
                    FreeGB  = $freeGB
                    FreePct = $pct
                }
            }
        } else {
            $lds = Get-CimInstance -CimSession $session -ClassName Win32_LogicalDisk -Filter 'DriveType = 3'
            $disks = foreach ($d in $lds) {
                $sizeGB = [math]::Round(($d.Size / 1GB), 2)
                $freeGB = [math]::Round(($d.FreeSpace / 1GB), 2)
                $pct    = if ($d.Size -gt 0) { [math]::Round(($d.FreeSpace / $d.Size) * 100, 2) } else { 0 }

                [pscustomobject]@{
                    Volume  = $d.DeviceID
                    SizeGB  = $sizeGB
                    FreeGB  = $freeGB
                    FreePct = $pct
                }
            }
        }

        return [pscustomobject]@{
            ComputerName = $ComputerName
            Error        = $null
            Disks        = $disks
            CimSession   = $session
        }
    } catch {
        return [pscustomobject]@{
            ComputerName = $ComputerName
            Error        = $_.Exception.Message
            Disks        = @()
            CimSession   = $session
        }
    }
    finally {
        if ($session) { Remove-CimSession $session -ErrorAction SilentlyContinue }
    }
}

function Write-Excel {
    param(
        [Parameter(Mandatory)] [array]$Data,
        [Parameter(Mandatory)] [string]$Path
    )

    if (-not $Data -or $Data.Count -eq 0) { return $false }

    # Try ImportExcel module
    try {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop

            $Data | Export-Excel -Path $Path -WorksheetName 'Disks' -AutoSize -BoldTopRow -FreezeTopRow -ClearSheet
            return $true
        }
    } catch { }

    # Try Excel COM automation
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)

        $headers = $Data[0].psobject.Properties.Name

        # Header
        for ($i=0; $i -lt $headers.Count; $i++) {
            $ws.Cells.Item(1, $i+1) = $headers[$i]
            $ws.Cells.Item(1, $i+1).Font.Bold = $true
        }

        # Rows
        for ($r=0; $r -lt $Data.Count; $r++) {
            foreach ($c in 0..($headers.Count - 1)) {
                $ws.Cells.Item($r+2, $c+1) = $Data[$r].$($headers[$c])
            }
        }

        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
        $wb.SaveAs($Path, 51)
        $wb.Close($true)
        $xl.Quit()
        return $true
    } catch { return $false }
}

function Write-HtmlReport {
    param(
        [Parameter(Mandatory)] [array]$Data,
        [Parameter(Mandatory)] [string]$Path,
        [int]$WarnGB,
        [int]$WarnPct,
        [switch]$ShowAll
    )

    $generated = Get-Date
    $title = "Server Low Disk Space Report"

    $subtitle = if ($ShowAll) {
        "All Disks (Low-Space Highlighted)"
    } else {
        "Low-Space Disks Only (≤ $WarnGB GB OR ≤ $WarnPct%)"
    }

    $totalServers = ($Data | Select-Object -ExpandProperty ServerName -Unique).Count
    $totalDisks   = $Data.Count
    $lowDisks     = ($Data | Where-Object { $_.FreeGB -le $WarnGB -or $_.FreePct -le $WarnPct }).Count

    $rows = foreach ($r in $Data) {
        $isLow = ($r.FreeGB -le $WarnGB -or $r.FreePct -le $WarnPct)
        $cls = if ($isLow) { "low" } else { "ok" }

        "<tr class='$cls'>
            <td>$($r.Domain)</td>
            <td>$($r.ServerName)</td>
            <td>$($r.IPv4)</td>
            <td>$($r.Volume)</td>
            <td class='num'>$($r.SizeGB)</td>
            <td class='num'>$($r.FreeGB)</td>
            <td class='num'>$($r.FreePct)</td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>$title</title>
<style>
body { font-family:Segoe UI,Tahoma,Arial; margin:20px; }
table { border-collapse:collapse; width:100%; }
th { background:#f1f1f1; position:sticky; top:0; padding:8px; }
td { padding:8px; border-bottom:1px solid #eee; }
tr.low { background:#fff5f5; }
.num { text-align:right; }
</style>
</head>
<body>

<h1>$title</h1>
<h3>$subtitle</h3>
<p>Generated: $generated</p>

<p><strong>Total Servers:</strong> $totalServers  
<strong>Total Disks:</strong> $totalDisks  
<strong>Low-Space Disks:</strong> $lowDisks</p>

<table>
<thead>
    <tr>
        <th>Domain</th>
        <th>Server</th>
        <th>IPv4</th>
        <th>Volume</th>
        <th>Size (GB)</th>
        <th>Free (GB)</th>
        <th>Free (%)</th>
    </tr>
</thead>
<tbody>
$($rows -join "`n")
</tbody>
</table>

</body>
</html>
"@

    $html | Out-File $Path -Encoding UTF8
}

# ----------------------------------------
# MAIN SCRIPT EXECUTION
# ----------------------------------------

Test-Module ActiveDirectory

$adParams = @{
    Filter         = 'OperatingSystem -like "*Server*"'
    Properties     = 'DNSHostName','DistinguishedName','enabled'
    ResultPageSize = 2000
}
if ($OU) { $adParams['SearchBase'] = $OU }

$servers = Get-ADComputer @adParams | Where-Object { $_.Enabled -eq $true }

$results = New-Object System.Collections.Generic.List[object]
$errors  = New-Object System.Collections.Generic.List[object]

$i = 0
$total = $servers.Count

foreach ($s in $servers) {
    $i++
    $name = $s.DNSHostName
    if (-not $name) { $name = $s.Name }
    $domain = Get-DomainFqdnFromDN $s.DistinguishedName

    Write-Progress -Activity "Gathering disk info" -Status "$name ($i of $total)" -PercentComplete (($i/$total)*100)

    $diskInfo = Get-ServerDisks -ComputerName $name -IncludeMountPoints:$IncludeMountPoints
    $ip = Resolve-IPv4 -ComputerName $name

    if ($diskInfo.Error) {
        $errors.Add([pscustomobject]@{ Computer=$name; Error=$diskInfo.Error })
        continue
    }

    foreach ($d in $diskInfo.Disks) {
        $results.Add([pscustomobject]@{
            Domain     = $domain
            ServerName = $name
            IPv4       = $ip
            Volume     = $d.Volume
            SizeGB     = $d.SizeGB
            FreeGB     = $d.FreeGB
            FreePct    = $d.FreePct
        })
    }
}

if (-not $ShowAll) {
    $results = $results | Where-Object { $_.FreeGB -le $WarningFreeGB -or $_.FreePct -le $WarningFreePct }
}

$results = $results | Sort-Object FreePct, FreeGB, ServerName

$results | Format-Table -AutoSize

# ----------------------------------------
# EXPORTS
# ----------------------------------------

$results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

Write-HtmlReport -Data $results -Path $HtmlPath -WarnGB $WarningFreeGB -WarnPct $WarningFreePct -ShowAll:$ShowAll

$ok = Write-Excel -Data $results -Path $XlsxPath

# ----------------------------------------
# COMMENTED-OUT EMAIL BLOCK
# ----------------------------------------

<#
# Only send if low disks found and ShowAll is NOT used
if (-not $ShowAll -and $results.Count -gt 0) {

    $MailParams = @{
        To      = 'DL-ServerOps@igt.com'   # <--- Change this
        From    = 'ServerReports@igt.com'  # <--- Change this
        Subject = "Low Disk Space Report - $($results.Count) Issues Detected"
        SmtpServer = 'SMTP.IGT.COM'
        Port    = 25
        UseSsl  = $false
        Body    = "Low disk space detected. Reports attached."
        Attachments = @($CsvPath, $XlsxPath, $HtmlPath)
    }

    try {
        Send-MailMessage @MailParams
        Write-Host "Email sent." -ForegroundColor Green
    }
    catch {
        Write-Warning "Email failed: $($_.Exception.Message)"
    }
}
#>

Write-Host "`nAll exports saved to: $ReportFolder" -ForegroundColor Cyan

if ($errors.Count -gt 0) {
    Write-Warning "`nSome servers failed:"
    $errors | Format-Table -AutoSize
}
