#Requires -Version 5.0

<#
.SYNOPSIS
    Analyzes Exchange SMTP Receive Protocol logs for mail traffic.

.DESCRIPTION
    Parses Exchange SMTP Receive Protocol log files (.log) and reports mail
    traffic per receive connector, including sender address, recipient(s),
    source IP, and connector name.

    A new log file is created every hour by Exchange. The script selects
    files based on their last-write time to cover the requested time window.

.PARAMETER LogPath
    Path to the folder containing SMTP Receive Protocol log files.
    Default: C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive

.PARAMETER Hours
    Number of hours back from now to include in the analysis. Default: 5.

.PARAMETER Connector
    Optional. Filter results to a specific connector name (partial match).

.PARAMETER ExportCsv
    Optional. Full path to a CSV file where results will be exported.
    If the file already exists it will be overwritten.

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -Hours 5

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -LogPath "D:\ExchangeLogs\SmtpReceive" -Hours 12

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -Hours 2 -Connector "Anon Relay"

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -Hours 24 -ExportCsv "C:\Reports\mail-traffic.csv"

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -Hours 5 -Connector "Anon Relay" -ExportCsv "C:\Reports\anon-relay.csv"
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$LogPath = 'C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive',

    [Parameter()]
    [ValidateRange(1, 8760)]
    [int]$Hours = 5,

    [Parameter()]
    [string]$Connector = '',

    [Parameter()]
    [string]$ExportCsv = ''
)

#region ── helpers ──────────────────────────────────────────────────────────────

function Get-EmailAddress {
    param([string]$Data, [string]$Command)
    # Match command like "MAIL From:<addr>" or "RCPT To:<addr>"
    if ($Data -match "(?i)^$Command[:\s]*<([^>]*)>") {
        return $Matches[1]
    }
    return $null
}

function Get-SourceIP {
    param([string]$RemoteEndpoint)
    # remote-endpoint format: ip:port  (IPv4) or [ipv6]:port
    if ($RemoteEndpoint -match '^(.+):(\d+)$') {
        return $Matches[1].Trim('[]')
    }
    return $RemoteEndpoint
}

function Get-ConnectorName {
    param([string]$ConnectorId)
    # connector-id format: SERVER\ConnectorName
    $idx = $ConnectorId.IndexOf('\')
    if ($idx -ge 0) {
        return $ConnectorId.Substring($idx + 1)
    }
    return $ConnectorId
}

#endregion

#region ── file selection ───────────────────────────────────────────────────────

$cutoff = (Get-Date).AddHours(-$Hours)

Write-Host "Exchange Receive Connector Traffic Analyzer" -ForegroundColor Cyan
Write-Host ("Analysing logs from: {0:yyyy-MM-dd HH:mm:ss}" -f $cutoff) -ForegroundColor Cyan
Write-Host ("Log path           : {0}" -f $LogPath) -ForegroundColor Cyan
Write-Host ""

if (-not (Test-Path $LogPath)) {
    Write-Error "Log path not found: $LogPath"
    exit 1
}

$logFiles = Get-ChildItem -Path $LogPath -Filter '*.log' -File |
    Where-Object { $_.LastWriteTime -ge $cutoff } |
    Sort-Object LastWriteTime

if ($logFiles.Count -eq 0) {
    Write-Warning "No log files found in the specified time window ($Hours hours)."
    exit 0
}

Write-Host ("Found {0} log file(s) to process." -f $logFiles.Count) -ForegroundColor Yellow
Write-Host ""

#endregion

#region ── parse logs ───────────────────────────────────────────────────────────

# Key: session-id  Value: hashtable with session data
$sessions = @{}

foreach ($file in $logFiles) {
    Write-Verbose "Processing: $($file.Name)"

    $lines = Get-Content -Path $file.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
    if (-not $lines) { continue }

    # Extract field names from the "#Fields:" header line
    $fieldsLine = $lines | Where-Object { $_ -like '#Fields:*' } | Select-Object -Last 1
    if (-not $fieldsLine) { continue }

    $fieldNames = ($fieldsLine -replace '^#Fields:\s*', '').Split(',')

    # Data lines are those that do NOT start with '#'
    $dataLines = $lines | Where-Object { $_ -notmatch '^#' -and $_.Trim() -ne '' }
    if (-not $dataLines) { continue }

    # Use ConvertFrom-Csv with the extracted field names
    $records = $dataLines | ConvertFrom-Csv -Header $fieldNames

    foreach ($rec in $records) {
        $sessionId  = $rec.'session-id'
        $event      = $rec.'event'
        $data       = $rec.'data'
        $connId     = $rec.'connector-id'
        $remoteEP   = $rec.'remote-endpoint'
        $timestamp  = $rec.'date-time'

        if (-not $sessionId) { continue }

        # Apply connector filter early
        if ($Connector -and $connId -notlike "*$Connector*") { continue }

        # Initialise session entry on first encounter
        if (-not $sessions.ContainsKey($sessionId)) {
            $sessions[$sessionId] = @{
                ConnectorId    = $connId
                ConnectorName  = Get-ConnectorName $connId
                SourceIP       = Get-SourceIP $remoteEP
                FirstSeen      = $timestamp
                MailFrom       = $null
                RcptTo         = [System.Collections.Generic.List[string]]::new()
                HasMail        = $false
            }
        }

        $s = $sessions[$sessionId]

        # Update source IP (in case first record had an empty remote-endpoint)
        if (-not $s.SourceIP -and $remoteEP) {
            $s.SourceIP = Get-SourceIP $remoteEP
        }

        # Only care about client-sent commands (event = "<")
        if ($event -ne '<') { continue }

        if ($data -like 'MAIL From:*' -or $data -like 'MAIL FROM:*') {
            $addr = Get-EmailAddress -Data $data -Command 'MAIL From'
            if ($addr -ne $null) {
                $s.MailFrom = $addr
                $s.HasMail  = $true
            }
        }
        elseif ($data -like 'RCPT To:*' -or $data -like 'RCPT TO:*') {
            $addr = Get-EmailAddress -Data $data -Command 'RCPT To'
            if ($addr -ne $null -and $s.RcptTo -notcontains $addr) {
                $s.RcptTo.Add($addr)
            }
        }
    }
}

#endregion

#region ── output results ───────────────────────────────────────────────────────

$results = $sessions.Values |
    Where-Object { $_.HasMail -eq $true -and $_.MailFrom -ne $null -and $_.RcptTo.Count -gt 0 } |
    ForEach-Object {
        foreach ($rcpt in $_.RcptTo) {
            [PSCustomObject]@{
                'Time (UTC)'    = $_.FirstSeen
                'Connector'     = $_.ConnectorName
                'Source IP'     = $_.SourceIP
                'From'          = $_.MailFrom
                'To'            = $rcpt
            }
        }
    } | Sort-Object 'Time (UTC)'

if ($results.Count -eq 0) {
    Write-Host "No mail traffic found in the last $Hours hour(s)." -ForegroundColor Yellow
}
else {
    Write-Host ("Found {0} message(s) across {1} session(s) in the last {2} hour(s):" -f
        $results.Count,
        ($sessions.Values | Where-Object HasMail).Count,
        $Hours) -ForegroundColor Green
    Write-Host ""
    $results | Format-Table -AutoSize -Wrap

    # Summary per connector
    Write-Host ""
    Write-Host "Summary per connector:" -ForegroundColor Cyan
    $results | Group-Object 'Connector' | ForEach-Object {
        Write-Host ("  {0,-45} : {1} message(s)" -f $_.Name, $_.Count)
    }

    # CSV export
    if ($ExportCsv) {
        try {
            $exportDir = Split-Path -Path $ExportCsv -Parent
            if ($exportDir -and -not (Test-Path $exportDir)) {
                New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
            }
            $results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Force
            Write-Host ""
            Write-Host ("Results exported to: {0}" -f $ExportCsv) -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to export CSV: $_"
        }
    }
}

#endregion
