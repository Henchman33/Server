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
    - Detects ImportExcel module and logs method used for XLSX export; falls back to Excel COM; or skips with guidance.

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

# Option C (custom path) — uncomment to use
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

Import-Module ImportExcel -ErrorAction SilentlyContinue
``

function Get-DomainFqdnFromDN {
    <#
        Takes a distinguishedName and returns the domain FQDN.
        Example: "CN=SERVER1,OU=Servers,DC=corp,DC=contoso,DC=com" -> "corp.contoso.com"
    #>
    param([Parameter(Mandatory)] [string]$DistinguishedName)

    if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return $null }

    (
        ($DistinguishedName -split ',' | ForEach-Object { $_.Trim() }) |
            Where-Object { $_ -like 'DC=*' } |
            ForEach-Object { $_.Substring(3) }
    ) -join '.'
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
        Exports to XLSX with clear detection/logging:
        - If ImportExcel is available: uses Export-Excel.
        - Else attempts Excel COM automation.
        - Else skips XLSX and returns guidance text.
        Returns: PSCustomObject { Success, Method, Message }
    #>
    param(
        [Parameter(Mandatory)] [array]$Data,
        [Parameter(Mandatory)] [string]$Path
    )

    $result = [pscustomobject]@{
        Success = $false
        Method  = $null
        Message = $null
    }

    if (-not $Data -or $Data.Count -eq 0) {
        $result.Message = "No data to export."
        return $result
    }

    # Detect ImportExcel
    $importExcelAvailable = $false
    try { $importExcelAvailable = bool } catch { $importExcelAvailable = $false }

    if ($importExcelAvailable) {
        try {
            Import-Module ImportExcel -ErrorAction Stop | Out-Null
            # Nicely formatted worksheet
            $Data | Export-Excel -Path $Path -WorksheetName 'Disks' -AutoSize -TableName 'DiskReport' -BoldTopRow -FreezeTopRow -ClearSheet
            $result.Success = $true
            $result.Method  = 'ImportExcel'
            $result.Message = 'Exported via ImportExcel module.'
            return $result
        } catch {
            $err = $_.Exception.Message
            Write-Warning "ImportExcel module is present but export failed: $err"
            Write-Host "Falling back to Excel COM automation..." -ForegroundColor Yellow
        }
    } else {
        Write-Host "ImportExcel module not found. Attempting Excel COM automation..." -ForegroundColor Yellow
        Write-Host "Tip: Install it with: Install-Module -Name ImportExcel -Scope CurrentUser" -ForegroundColor DarkYellow
    }

    # Try Excel COM
    $xl = $null; $wb = $null; $ws = $null
    try {
        $xl = New-Object -ComObject Excel.Application
    } catch {
        $result.Message = "Excel COM automation is not available. To enable .xlsx export, either install ImportExcel (`Install-Module ImportExcel -Scope CurrentUser`) or install Microsoft Excel on this machine."
        return $result
    }

    try {
        $xl.Visible = $false
        $wb = $xl.Workbooks.Add()
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = 'Disks'

        $headers = $Data[0].psobject.Properties.Name

        # Header row
        for ($i=0; $i -lt $headers.Count; $i++) {
            $ws.Cells.Item(1, $i+1) = $headers[$i]
            $ws.Cells.Item(1, $i+1).Font.Bold = $true
        }

        # Data rows
        for ($r=0; $r -lt $Data.Count; $r++) {
            for ($c=0; $c -lt $headers.Count; $c++) {
                $ws.Cells.Item($r+2, $c+1) = $Data[$r].$($headers[$c])
            }
        }

        # Autofit columns
        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null

        # Save as XLSX (file format 51)
        $wb.SaveAs($Path, 51)
        $wb.Close($true)
        $xl.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)

        $result.Success = $true
        $result.Method  = 'ExcelCOM'
        $result.Message = 'Exported via Excel COM automation.'
        return $result
    } catch {
        try { if ($wb) { $wb.Close($false) } } catch {}
        try { if ($xl) { $xl.Quit() } } catch {}
        $result.Message = "Excel COM export failed: $($_.Exception.Message)"
        return $result
    }
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
body { font-family:Segoe UI,Tahoma,Arial; margin:20px; color:#1b1b1b;}
table { border-collapse:collapse; width:100%; }
th { background:#f1f1f1; position:sticky; top:0; padding:8px; text-align:left; }
td { padding:8px; border-bottom:1px solid #eee; }
tr.low { background:#fff5f5; }
.num { text-align:right; white-space:nowrap; }
.summary span { display:inline-block; margin-right:14px; background:#f7f7f7; padding:6px 10px; border-radius:6px;}
.search { margin: 12px 0; }
input[type="search"] { padding:6px 10px; width:320px; border:1px solid #ccc; border-radius:4px; }
</style>
<script>
function filterTable() {
    const q = document.getElementById('q').value.toLowerCase();
    const rows = document.querySelectorAll('#data tbody tr');
    rows.forEach(r => { r.style.display = r.textContent.toLowerCase().includes(q) ? '' : 'none'; });
}
</script>
</head>
<body>

<h1>$title</h1>
<h3>$subtitle</h3>
<div class="summary">
  <span><strong>Generated:</strong> $generated</span>
  <span><strong>Total Servers:</strong> $totalServers</span>
  <span><strong>Total Disks:</strong> $totalDisks</span>
  <span><strong>Low-Space Disks:</strong> $lowDisks</span>
  <span><strong>Thresholds:</strong> ≤ $WarnGB GB OR ≤ $WarnPct%</span>
</div>

<div class="search">
  <label for="q"><strong>Filter:</strong></label>
  <input id="q" type="search" placeholder="Type to filter (server, IP, drive, etc.)" oninput="filterTable()">
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

try {
    Test-Module -Name ActiveDirectory
} catch {
    Write-Error $_.Exception.Message
    break
}

$adParams = @{
    Filter         = 'OperatingSystem -like "*Server*"'
    Properties     = @('DNSHostName','DistinguishedName','enabled')
    ResultPageSize = 2000
    ResultSetSize  = $null
}
if ($OU) { $adParams['SearchBase'] = $OU }

$servers = Get-ADComputer @adParams | Where-Object { $_.Enabled -ne $false } | Sort-Object -Property Name

if (-not $servers) {
    Write-Warning "No server accounts found in AD with the given criteria."
    return
}

$results = New-Object System.Collections.Generic.List[object]
$errors  = New-Object System.Collections.Generic.List[object]

$i = 0
$total = $servers.Count

foreach ($s in $servers) {
    $i++
    $name = if ($s.DNSHostName) { $s.DNSHostName } else { $s.Name }
    $domain = Get-DomainFqdnFromDN -DistinguishedName $s.DistinguishedName

    Write-Progress -Activity "Gathering disk info" -Status "$name ($i of $total)" -PercentComplete (($i/$total)*100)

    $diskInfo = Get-ServerDisks -ComputerName $name -IncludeMountPoints:$IncludeMountPoints
    $ip = Resolve-IPv4 -ComputerName $name

    if ($diskInfo.Error) {
        $errors.Add([pscustomobject]@{ Computer=$name; Error=$diskInfo.Error }) | Out-Null
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
        }) | Out-Null
    }
}

# Filter unless showing all
if (-not $ShowAll) {
    $results = $results | Where-Object { $_.FreeGB -le $WarningFreeGB -or $_.FreePct -le $WarningFreePct }
}

# Sort for readability
$results = $results | Sort-Object -Property @{Expression='FreePct'; Ascending=$true}, @{Expression='FreeGB'; Ascending=$true}, ServerName, Volume

# Output to console
$results | Format-Table -AutoSize Domain, ServerName, IPv4, Volume, SizeGB, FreeGB, FreePct

# ----------------------------------------
# EXPORTS
# ----------------------------------------

# CSV
try {
    $null = $results | Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8
    Write-Host "CSV exported to: $CsvPath" -ForegroundColor Green
} catch {
    Write-Warning "Failed to export CSV: $($_.Exception.Message)"
}

# HTML
try {
    Write-HtmlReport -Data $results -Path $HtmlPath -WarnGB $WarningFreeGB -WarnPct $WarningFreePct -ShowAll:$ShowAll
    Write-Host "HTML report exported to: $HtmlPath" -ForegroundColor Green
} catch {
    Write-Warning "Failed to write HTML: $($_.Exception.Message)"
}

# XLSX with detection & logging
$xlResult = Write-Excel -Data $results -Path $XlsxPath
if ($xlResult.Success) {
    Write-Host "XLSX exported to: $XlsxPath (method: $($xlResult.Method))" -ForegroundColor Green
} else {
    Write-Warning "XLSX export skipped: $($xlResult.Message)"
}

# ----------------------------------------
# OPTIONAL: Commented-out email section
# ----------------------------------------
<#
# Only send if low disks found and -ShowAll is NOT used
if (-not $ShowAll -and $results.Count -gt 0) {

    $MailParams = @{
        To      = 'DL-ServerOps@igt.com'       # <-- Change to your distro/team
        From    = 'ServerReports@igt.com'      # <-- Change as needed
        Subject = ("LOW DISK SPACE: {0} disks across {1} servers (≤ {2} GB OR ≤ {3}%)" -f `
                    $results.Count, `
                    ($results | Select-Object -ExpandProperty ServerName -Unique).Count, `
                    $WarningFreeGB, $WarningFreePct)
        SmtpServer = 'SMTP.IGT.COM'            # <-- Your SMTP server
        Port    = 25                            # <-- Change if needed (e.g., 587)
        UseSsl  = $false                        # <-- Set to $true if SMTP requires TLS
        # Credential = (Get-Credential)         # <-- Uncomment if SMTP requires auth
        Body    = @"
Hi team,

Attached are the latest low disk space reports:

• CSV
• XLSX
• HTML (open in a browser for an interactive, filterable view)

Thresholds used: ≤ $WarningFreeGB GB OR ≤ $WarningFreePct%
Generated on: $(Get-Date)

Summary:
- Total servers impacted: $(( $results | Select-Object -ExpandProperty ServerName -Unique ).Count)
- Low-space disks: $($results.Count)

Regards,
Automation
"@
        Attachments = @($CsvPath, $XlsxPath, $HtmlPath)    # Attaches all three reports
        BodyAsHtml  = $false
        DeliveryNotificationOption = 'OnFailure','OnSuccess'
    }

    try {
        Send-MailMessage @MailParams
        Write-Host "Notification email sent via $($MailParams.SmtpServer)" -ForegroundColor Green
    } catch {
        Write-Warning "Failed to send email: $($_.Exception.Message)"
    }
} else {
    Write-Host "No email sent (either -ShowAll used or no low-space disks found)." -ForegroundColor Yellow
}
#>

# ----------------------------------------
# Errors & Footer
# ----------------------------------------

if ($errors.Count -gt 0) {
    Write-Warning "`nSome servers could not be queried:"
    $errors | Format-Table -AutoSize
}

Write-Host "`nAll requested exports saved under: $ReportFolder" -ForegroundColor Cyan
