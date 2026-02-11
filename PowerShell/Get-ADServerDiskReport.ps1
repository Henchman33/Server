<#
.SYNOPSIS
    Enumerates AD Windows Server computers and reports disk usage to CSV, XLSX, and HTML.

.DESCRIPTION
    - Finds server-class computers in Active Directory.
    - Collects domain (FQDN), IPv4, and disk info for each fixed drive (optionally volumes/mount points).
    - Flags low free space by GB and/or % thresholds.
    - Outputs to screen and exports CSV, XLSX, and a styled HTML report into a Desktop folder:
        "Server Low Disk Space Report" (created if missing).

.PARAMETER WarningFreeGB
    Low-space threshold in GB. Default: 20

.PARAMETER WarningFreePct
    Low-space threshold in percent. Default: 10

.PARAMETER ShowAll
    If set, includes all disks (not only low-space ones).

.PARAMETER IncludeMountPoints
    If set, includes NTFS mount points/volumes without drive letters (Win32_Volume).

.PARAMETER OU
    Optional DN to scope AD search (e.g., "OU=Servers,DC=corp,DC=contoso,DC=com").

.NOTES
    - Run in Windows PowerShell 5.1 on a DC/admin host with ActiveDirectory module available.
    - Requires admin rights on target servers for WMI/CIM (WSMan with fallback to DCOM).
    - XLSX export:
        * Prefers ImportExcel module if installed (Export-Excel).
        * Falls back to Excel COM automation if Microsoft Excel is installed.
        * If neither is available, CSV and HTML are produced and a warning is shown.
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

# Option A (default): current user's Desktop\Server Low Disk Space Report
$ReportFolder = Join-Path -Path ([Environment]::GetFolderPath('Desktop')) -ChildPath 'Server Low Disk Space Report'

# Option B (custom path): Uncomment and adjust as needed
# $ReportFolder = 'D:\Reports\Server Low Disk Space Report'

# Ensure folder exists
if (-not (Test-Path -Path $ReportFolder)) {
    New-Item -Path $ReportFolder -ItemType Directory -Force | Out-Null
}

# Filenames with timestamp
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$CsvPath  = Join-Path $ReportFolder "ServerDiskReport_$stamp.csv"
$XlsxPath = Join-Path $ReportFolder "ServerDiskReport_$stamp.xlsx"
$HtmlPath = Join-Path $ReportFolder "ServerDiskReport_$stamp.html"

# -------------------------
# Helper functions
# -------------------------

function Test-Module {
    param([Parameter(Mandatory)] [string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' is not available. Install RSAT/AD tools and try again."
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
            $ip = Get-CimInstance -CimSession $CimSession -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled = True' -ErrorAction Stop |
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
    <#
        Returns: [pscustomobject] with:
            ComputerName, Error, Disks (list of Volume, SizeGB, FreeGB, FreePct), CimSession
    #>
    param(
        [Parameter(Mandatory)] [string]$ComputerName,
        [switch]$IncludeMountPoints
    )

    $session = New-BestCimSession -ComputerName $ComputerName
    if (-not $session) {
        return [pscustomobject]@{
            ComputerName = $ComputerName
            Error        = "Unable to create CIM session (WSMan & DCOM failed)"
            Disks        = @()
            CimSession   = $null
        }
    }

    try {
        if ($IncludeMountPoints) {
            $vols = Get-CimInstance -CimSession $session -ClassName Win32_Volume -Filter 'DriveType = 3' -ErrorAction Stop |
                Where-Object { $_.Capacity -gt 0 }
            $disks = foreach ($v in $vols) {
                $sizeGB = [math]::Round(($v.Capacity / 1GB), 2)
                $freeGB = [math]::Round(($v.FreeSpace / 1GB), 2)
                $pct    = if ($v.Capacity -gt 0) { [math]::Round((($v.FreeSpace / $v.Capacity) * 100), 2) } else { 0 }
                $name   = if ($v.DriveLetter) { $v.DriveLetter } elseif ($v.Name) { $v.Name.TrimEnd('\') } else { $v.Label }
                [pscustomobject]@{
                    Volume  = $name
                    SizeGB  = $sizeGB
                    FreeGB  = $freeGB
                    FreePct = $pct
                }
            }
        } else {
            $lds = Get-CimInstance -CimSession $session -ClassName Win32_LogicalDisk -Filter 'DriveType = 3' -ErrorAction Stop
            $disks = foreach ($d in $lds) {
                $sizeGB = [math]::Round(($d.Size / 1GB), 2)
                $freeGB = [math]::Round(($d.FreeSpace / 1GB), 2)
                $pct    = if ($d.Size -gt 0) { [math]::Round((($d.FreeSpace / $d.Size) * 100), 2) } else { 0 }
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
        if ($session) { Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue }
    }
}

function Write-Excel {
    <#
        Best-effort XLSX export.
        Prefers ImportExcel (Export-Excel). Falls back to Excel COM. Else returns $false.
    #>
    param(
        [Parameter(Mandatory)] [array]$Data,
        [Parameter(Mandatory)] [string]$Path
    )

    if (-not $Data -or $Data.Count -eq 0) {
        Write-Warning "No data to export to Excel."
        return $false
    }

    # Try ImportExcel
    try {
        if (Get-Module -ListAvailable -Name ImportExcel) {
            Import-Module ImportExcel -ErrorAction Stop | Out-Null
            # Auto-size columns, add a table and freeze header row
            $Data | Export-Excel -Path $Path -WorksheetName 'Disks' -AutoSize -TableName 'DiskReport' -BoldTopRow -FreezeTopRow -ClearSheet
            # Add conditional formatting for FreePct and FreeGB
            $excelPkg = Open-ExcelPackage -Path $Path
            $ws = $excelPkg.Workbook.Worksheets['Disks']
            # Find column indexes
            $headerMap = @{}
            for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
                $headerMap[$ws.Cells[1,$c].Text] = $c
            }
            $rowStart = 2
            $rowEnd   = $ws.Dimension.End.Row
            if ($headerMap['FreePct']) {
                Add-ConditionalFormatting -Worksheet $ws -Address ($ws.Cells[$rowStart,$headerMap['FreePct'],$rowEnd,$headerMap['FreePct']].Address) -RuleType ThreeColorScale | Out-Null
            }
            if ($headerMap['FreeGB']) {
                Add-ConditionalFormatting -Worksheet $ws -Address ($ws.Cells[$rowStart,$headerMap['FreeGB'],$rowEnd,$headerMap['FreeGB']].Address) -RuleType ThreeColorScale | Out-Null
            }
            Close-ExcelPackage $excelPkg -Show:$false
            return $true
        }
    } catch {
        Write-Verbose "ImportExcel export failed: $($_.Exception.Message)"
    }

    # Try Excel COM (requires Excel to be installed)
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = 'Disks'

        # Write header
        $headers = $Data[0].psobject.Properties.Name
        for ($i=0; $i -lt $headers.Count; $i++) {
            $ws.Cells.Item(1, $i+1) = $headers[$i]
            $ws.Cells.Item(1, $i+1).Font.Bold = $true
        }

        # Write rows
        for ($r=0; $r -lt $Data.Count; $r++) {
            $row = $Data[$r]
            for ($c=0; $c -lt $headers.Count; $c++) {
                $ws.Cells.Item($r+2, $c+1) = $row.$($headers[$c])
            }
        }

        # Auto-fit
        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

        # Save as XLSX (FileFormat = 51)
        $xlFileFormat = 51
        $wb.SaveAs($Path, $xlFileFormat)
        $wb.Close($true)
        $xl.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl) | Out-Null
        return $true
    } catch {
        Write-Verbose "Excel COM export failed: $($_.Exception.Message)"
        try { if ($wb) { $wb.Close($false) } } catch {}
        try { if ($xl) { $xl.Quit() } } catch {}
        return $false
    }
}

function Write-HtmlReport {
    param(
        [Parameter(Mandatory)] [array]$Data,
        [Parameter(Mandatory)] [string]$Path,
        [int]$WarnGB = 20,
        [int]$WarnPct = 10,
        [switch]$ShowAll
    )

    $generated = Get-Date
    $title = "Server Low Disk Space Report"
    $subtitle = if ($ShowAll) {
        "All disks (highlighting low space)"
    } else {
        "Only low-space disks (≤ $WarnGB GB OR ≤ $WarnPct%)"
    }

    # Build rows with CSS class based on thresholds
    $rows = foreach ($r in $Data) {
        $isLow = ($r.FreeGB -le $WarnGB -or $r.FreePct -le $WarnPct)
        $cls = if ($isLow) { 'low' } else { 'ok' }
        "<tr class='$cls'><td>$($r.Domain)</td><td>$($r.ServerName)</td><td>$($r.IPv4)</td><td>$($r.Volume)</td><td class='num'>$($r.SizeGB)</td><td class='num'>$($r.FreeGB)</td><td class='num'>$($r.FreePct)</td></tr>"
    }

    $totalServers = ($Data | Select-Object -ExpandProperty ServerName -Unique).Count
    $totalDisks   = $Data.Count
    $lowDisks     = ($Data | Where-Object { $_.FreeGB -le $WarnGB -or $_.FreePct -le $WarnPct }).Count

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>$title</title>
<style>
    body { font-family: Segoe UI, Tahoma, Arial, sans-serif; color:#1b1b1b; margin:20px; }
    h1 { margin:0 0 6px 0; }
    h2 { margin:0 0 16px 0; font-weight:normal; color:#555; }
    .meta { margin: 8px 0 16px 0; color:#666; }
    .summary { margin: 12px 0 16px 0; }
    .summary span { display:inline-block; margin-right:18px; padding:6px 10px; background:#f1f1f1; border-radius:6px; }
    .legend { margin: 8px 0 18px 0; color:#555; }
    table { border-collapse: collapse; width:100%; }
    thead th { position: sticky; top: 0; background: #fafafa; border-bottom: 2px solid #ddd; text-align:left; padding:8px; }
    tbody td { border-bottom: 1px solid #eee; padding:8px; }
    tbody tr:nth-child(even) { background: #fcfcfc; }
    tbody tr.low { background: #fff5f5; }
    tbody tr.low td { border-bottom-color:#f2dcdc; }
    .num { text-align:right; white-space:nowrap; }
    .foot { margin-top:16px; font-size:12px; color:#777; }
    .search { margin: 12px 0; }
    input[type="search"] { padding:6px 10px; width:320px; border:1px solid #ccc; border-radius:4px; }
</style>
<script>
// Simple client-side filter
function filterTable() {
    const q = document.getElementById('q').value.toLowerCase();
    const rows = document.querySelectorAll('#data tbody tr');
    rows.forEach(r => {
        r.style.display = r.textContent.toLowerCase().includes(q) ? '' : 'none';
    });
}
</script>
</head>
<body>
<h1>$title</h1>
<h2>$subtitle</h2>
<div class="meta">Generated: $generated</div>
<div class="summary">
    <span><strong>Total servers:</strong> $totalServers</span>
    <span><strong>Total disks:</strong> $totalDisks</span>
    <span><strong>Low-space disks:</strong> $lowDisks</span>
    <span><strong>Thresholds:</strong> ≤ $WarnGB GB OR ≤ $WarnPct%</span>
</div>
<div class="legend">Rows shaded light red are below threshold.</div>

<div class="search">
    <label for="q"><strong>Filter:</strong></label>
    <input id="q" type="search" placeholder="Type to filter by server, drive, IP, etc." oninput="filterTable()">
</div>

<table id="data">
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
        $($rows -join "`n        ")
    </tbody>
</table>

<div class="foot">Report saved to: $Path</div>
</body>
</html>
"@

    $html | Out-File -FilePath $Path -Encoding UTF8
}

# -------------------------
# Main
# -------------------------

try {
    Test-Module -Name ActiveDirectory
} catch {
    Write-Error $_.Exception.Message
    break
}

Write-Verbose "Querying Active Directory for server computer accounts..."

$adParams = @{
    Filter           = 'OperatingSystem -like "*Server*"'
    Properties       = @('DNSHostName','DistinguishedName','OperatingSystem','enabled')
    ResultPageSize   = 2000
    ResultSetSize    = $null
}
if ($OU) { $adParams['SearchBase'] = $OU }

$servers = Get-ADComputer @adParams | Where-Object { $_.Enabled -ne $false } | Sort-Object -Property Name

if (-not $servers) {
    Write-Warning "No server accounts found in AD with the given criteria."
    return
}

$results = New-Object System.Collections.Generic.List[object]
$errors  = New-Object System.Collections.Generic.List[object]

$idx = 0
$total = $servers.Count

foreach ($s in $servers) {
    $idx++
    $name   = if ($s.DNSHostName) { $s.DNSHostName } else { $s.Name }
    $domain = Get-DomainFqdnFromDN -DistinguishedName $s.DistinguishedName

    Write-Progress -Activity "Collecting disk info" -Status "$name ($idx of $total)" -PercentComplete (($idx / $total) * 100)

    $diskInfo = Get-ServerDisks -ComputerName $name -IncludeMountPoints:$IncludeMountPoints
    $ipv4 = Resolve-IPv4 -ComputerName $name

    if ($diskInfo.Error) {
        $errors.Add([pscustomobject]@{
            Computer = $name
            Error    = $diskInfo.Error
        }) | Out-Null
        continue
    }

    foreach ($d in $diskInfo.Disks) {
        $results.Add([pscustomobject]@{
            Domain     = $domain
            ServerName = $name
            IPv4       = $ipv4
            Volume     = $d.Volume
            SizeGB     = $d.SizeGB
            FreeGB     = $d.FreeGB
            FreePct    = $d.FreePct
        }) | Out-Null
    }
}

# Filter if not showing all
if (-not $ShowAll) {
    $results = $results | Where-Object { $_.FreeGB -le $WarningFreeGB -or $_.FreePct -le $WarningFreePct }
}

# Sort for readability
$results = $results | Sort-Object -Property @{Expression='FreePct'; Ascending=$true}, @{Expression='FreeGB'; Ascending=$true}, ServerName, Volume

# Output to console
$results | Format-Table -AutoSize Domain, ServerName, IPv4, Volume, SizeGB, FreeGB, FreePct

# -------------------------
# Exports
# -------------------------

# CSV
try {
    $null = $results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvPath
    Write-Host "CSV exported to: $CsvPath" -ForegroundColor Green
} catch {
    Write-Warning "Failed to export CSV: $($_.Exception.Message)"
}

# XLSX (best-effort)
$xlsxOK = $false
try {
    $xlsxOK = Write-Excel -Data $results -Path $XlsxPath
    if ($xlsxOK) {
        Write-Host "XLSX exported to: $XlsxPath" -ForegroundColor Green
    } else {
        Write-Warning "XLSX export was not completed (ImportExcel/Excel not available). CSV and HTML were still generated."
    }
} catch {
    Write-Warning "XLSX export failed: $($_.Exception.Message)"
}

# HTML
try {
    Write-HtmlReport -Data $results -Path $HtmlPath -WarnGB $WarningFreeGB -WarnPct $WarningFreePct -ShowAll:$ShowAll
    Write-Host "HTML report exported to: $HtmlPath" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write HTML: $($_.Exception.Message)"
}

# Errors (if any)
if ($errors.Count -gt 0) {
    Write-Warning "`nSome servers could not be queried:"
    $errors | Format-Table -AutoSize
}

Write-Host "`nAll requested exports written under: $ReportFolder" -ForegroundColor Cyan
