# ================================
# System Event Summary for DC
# ================================

Write-Host "Collecting system event information..." -ForegroundColor Cyan

# 1. Get the most recent System event (by EventRecordID)
$LastSystemEvent = Get-WinEvent -LogName System -MaxEvents 1

# 2. Get last reboot time (reliable)
$LastBootTime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime

# 3. Find last reboot/shutdown event with user (Event ID 1074)
$LastRebootEvent = Get-WinEvent -FilterHashtable @{
    LogName = 'System'
    Id      = 1074
} -MaxEvents 1 -ErrorAction SilentlyContinue

# 4. Output results
Write-Host "`n=== System Event Information ===" -ForegroundColor Yellow

[PSCustomObject]@{
    ComputerName        = $env:COMPUTERNAME
    LastSystemEventID   = $LastSystemEvent.Id
    LastEventRecordID   = $LastSystemEvent.RecordId
    LastEventTime       = $LastSystemEvent.TimeCreated
    LastBootTime        = $LastBootTime
    RebootInitiatedBy   = if ($LastRebootEvent) { $LastRebootEvent.Properties[6].Value } else { "Unknown / Not recorded" }
    RebootReason        = if ($LastRebootEvent) { $LastRebootEvent.Properties[2].Value } else { "Unexpected reboot or power loss" }
    RebootProcess       = if ($LastRebootEvent) { $LastRebootEvent.Properties[0].Value } else { "N/A" }
} | Format-List
