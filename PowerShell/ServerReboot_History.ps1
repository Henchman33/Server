<#
.SYNOPSIS
    Retrieves the reboot history of a local or remote server to find out who restarted it and why.
.PARAMETER ComputerName
    The name of the computer to query. Defaults to the local machine.
.PARAMETER Days
    How many days of history to look back. Defaults to 7 days.
#>
param(
    [string]$ComputerName = "localhost",
    [int]$Days = 7
)

# Calculate the start time for logs
$StartTime = (Get-Date).AddDays(-$Days)

Write-Host "Fetching reboot logs from '$ComputerName' for the last $Days days..." -ForegroundColor Cyan

try {
    # Define the Event IDs we care about
    # 1074 = Normal Shutdown/Restart, 6008 = Unexpected Shutdown
    $FilterHashtable = @{
        LogName   = 'System'
        Id        = 1074, 6008
        StartTime = $StartTime
    }

    # Fetch events
    $Events = Get-WinEvent -ComputerName $ComputerName -FilterHashtable $FilterHashtable -ErrorAction Stop
    
    $RebootHistory = foreach ($Event in $Events) {
        if ($Event.Id -eq 1074) {
            # Parse the XML data for Event 1074 to extract specific attributes
            $EventXML = [xml]$Event.ToXml()
            $Data = $EventXML.Event.EventData.Data

            [PSCustomObject]@{
                Timestamp    = $Event.TimeCreated
                ComputerName = $ComputerName
                EventID      = $Event.Id
                Type         = "Graceful/Planned"
                Action       = $Data[4].'#text' # Shutdown or Restart
                User         = $Data[6].'#text' # User who initiated it
                Process      = $Data[0].'#text' # Process that called it (e.g. winlogon, Explorer)
                Reason       = $Data[2].'#text' # Comment or reason code provided
            }
        }
        elseif ($Event.Id -eq 6008) {
            # Event 6008 details are harder to parse directly into words, as it implies a sudden crash
            [PSCustomObject]@{
                Timestamp    = $Event.TimeCreated
                ComputerName = $ComputerName
                EventID      = $Event.Id
                Type         = "UNEXPECTED CRASH / POWER LOSS"
                Action       = "N/A"
                User         = "N/A"
                Process      = "System Crash"
                Reason       = "The previous system shutdown was unexpected."
            }
        }
    }

    # Output formatted results to the console
    if ($RebootHistory) {
        $RebootHistory | Sort-Object Timestamp -Descending | Format-Table -AutoSize
    } else {
        Write-Host "No reboot events found within the last $Days days." -ForegroundColor Yellow
    }

}
catch {
    Write-Error "Failed to retrieve logs from $ComputerName. Reason: $_"
}
