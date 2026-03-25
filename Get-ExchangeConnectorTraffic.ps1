#Requires -Version 5.0

<#
.SYNOPSIS
    Analyzes Exchange SMTP Receive connector protocol logs for mail traffic.

.DESCRIPTION
    Parses Exchange SMTP Receive Protocol log files (.log) and reports mail
    traffic per receive connector, including sender address, recipient(s),
    source IP, and connector name.

    A new log file is created every hour by Exchange. Files are selected based
    on their last-write time to cover the requested time window.

    An executive summary (per server -> per connector -> per source IP with
    message counts) is always shown on the console. When -ExportCsv is used,
    the summary is also saved as a .txt file with the same base name.

.PARAMETER LogPath
    Path to the folder containing SMTP Receive Protocol log files.
    Default: C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive
    Override if logs are stored on a custom path.

.PARAMETER Hours
    Number of hours back from now to analyse. Cannot be combined with -Days.
    If neither -Hours nor -Days is specified, defaults to 5 hours.

.PARAMETER Days
    Number of days back from now to analyse. Cannot be combined with -Hours.

.PARAMETER Connector
    Optional. Filter results to a specific connector name (partial, case-insensitive match).

.PARAMETER ExcludeIP
    Optional. One or more IP addresses to exclude from the report.
    Useful for filtering out known infrastructure (e.g. Edge Transport servers)
    whose traffic is expected and not relevant to the analysis.

.PARAMETER ExportCsv
    Optional. Full path to a CSV file for detailed results.
    A matching summary text file (<basename>_summary.txt) is created automatically.
    Both files are overwritten if they already exist.

.EXAMPLE
    .\Get-ExchangeConnectorTraffic.ps1 -LogPath "C:\Logs\SmtpReceive"

.EXAMPLE
    .\Get-ExchangeConnectorTraffic.ps1 -LogPath "C:\Logs\SmtpReceive" -Hours 12

.EXAMPLE
    .\Get-ExchangeConnectorTraffic.ps1 -LogPath "C:\Logs\SmtpReceive" -Days 3

.EXAMPLE
    .\Get-ExchangeConnectorTraffic.ps1 -LogPath "C:\Logs\SmtpReceive" -Hours 5 -Connector "Anon Relay"

.EXAMPLE
    .\Get-ExchangeConnectorTraffic.ps1 -LogPath "C:\Logs\SmtpReceive" -Days 7 -ExportCsv "C:\Reports\traffic.csv"

.EXAMPLE
    .\Get-ExchangeConnectorTraffic.ps1 -LogPath "C:\Logs\SmtpReceive" -Hours 5 -ExcludeIP "10.116.1.10","10.116.1.11"
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$LogPath = 'C:\Program Files\Microsoft\Exchange Server\V15\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive',

    [Parameter()]
    [ValidateRange(1, 8760)]
    [int]$Hours = 0,

    [Parameter()]
    [ValidateRange(1, 365)]
    [int]$Days = 0,

    [Parameter()]
    [string]$Connector = '',

    [Parameter()]
    [string[]]$ExcludeIP = @(),

    [Parameter()]
    [string]$ExportCsv = ''
)

#region ── validation ───────────────────────────────────────────────────────────

if ($Hours -gt 0 -and $Days -gt 0) {
    Write-Error "Specify either -Hours or -Days, not both."
    exit 1
}

if ($Days -gt 0) {
    $cutoff     = (Get-Date).AddDays(-$Days)
    $windowDesc = "$Days day(s)"
}
elseif ($Hours -gt 0) {
    $cutoff     = (Get-Date).AddHours(-$Hours)
    $windowDesc = "$Hours hour(s)"
}
else {
    $cutoff     = (Get-Date).AddHours(-5)
    $windowDesc = '5 hours (default)'
}

#endregion

#region ── helpers ──────────────────────────────────────────────────────────────

function Get-EmailAddress {
    param([string]$Data, [string]$Command)
    if ($Data -match "(?i)^$Command[:\s]*<([^>]*)>") {
        return $Matches[1]
    }
    return $null
}

function Get-SourceIP {
    param([string]$Endpoint)
    # Endpoint format: ip:port (IPv4) or [ipv6]:port
    if ($Endpoint -match '^(.+):\d+$') {
        return $Matches[1].Trim('[]')
    }
    return $Endpoint
}

function Get-ServerName {
    param([string]$ConnectorId)
    $idx = $ConnectorId.IndexOf('\')
    if ($idx -ge 0) { return $ConnectorId.Substring(0, $idx) }
    return $ConnectorId
}

function Get-ConnectorName {
    param([string]$ConnectorId)
    $idx = $ConnectorId.IndexOf('\')
    if ($idx -ge 0) { return $ConnectorId.Substring($idx + 1) }
    return $ConnectorId
}

#endregion

#region ── file selection ───────────────────────────────────────────────────────

Write-Host "Exchange Receive Connector Traffic Analyzer" -ForegroundColor Cyan
Write-Host ("Period             : last {0}" -f $windowDesc) -ForegroundColor Cyan
Write-Host ("Analyse from (UTC) : {0:yyyy-MM-dd HH:mm:ss}" -f $cutoff.ToUniversalTime()) -ForegroundColor Cyan
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
    Write-Warning "No log files found in the specified time window ($windowDesc)."
    exit 0
}

Write-Host ("Found {0} log file(s) to process." -f $logFiles.Count) -ForegroundColor Yellow
Write-Host ""

#endregion

#region ── parse logs ───────────────────────────────────────────────────────────

# Key: session-id  Value: hashtable with session data
$sessions  = @{}
$fileIndex = 0
$totalFiles = $logFiles.Count

foreach ($file in $logFiles) {
    $fileIndex++
    $pct = [int]($fileIndex / $totalFiles * 100)

    Write-Progress -Activity "Processing log files" `
                   -Status ("File {0}/{1}  ({2})  -  {3} session(s) found so far" -f
                       $fileIndex, $totalFiles, $file.Name, $sessions.Count) `
                   -PercentComplete $pct

    $lines = Get-Content -Path $file.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
    if (-not $lines) { continue }

    # Extract field names from the "#Fields:" header line
    $fieldsLine = $lines | Where-Object { $_ -like '#Fields:*' } | Select-Object -Last 1
    if (-not $fieldsLine) { continue }

    $fieldNames = ($fieldsLine -replace '^#Fields:\s*', '').Split(',')

    # Data lines are those that do NOT start with '#'
    $dataLines = $lines | Where-Object { $_ -notmatch '^#' -and $_.Trim() -ne '' }
    if (-not $dataLines) { continue }

    $records = $dataLines | ConvertFrom-Csv -Header $fieldNames

    foreach ($rec in $records) {
        $sessionId = $rec.'session-id'
        $event     = $rec.'event'
        $data      = $rec.'data'
        $connId    = $rec.'connector-id'
        $remoteEP  = $rec.'remote-endpoint'
        $timestamp = $rec.'date-time'

        if (-not $sessionId) { continue }

        # Apply connector filter early
        if ($Connector -and $connId -notlike "*$Connector*") { continue }

        # Initialise session entry on first encounter
        if (-not $sessions.ContainsKey($sessionId)) {
            $sessions[$sessionId] = @{
                ServerName    = Get-ServerName    $connId
                ConnectorName = Get-ConnectorName $connId
                SourceIP      = Get-SourceIP      $remoteEP
                FirstSeen     = $timestamp
                MailFrom      = $null
                RcptTo        = [System.Collections.Generic.List[string]]::new()
                HasMail       = $false
            }
        }

        $s = $sessions[$sessionId]

        # Capture source IP from the first non-empty remote endpoint seen for this session
        if (-not $s.SourceIP -and $remoteEP) {
            $s.SourceIP = Get-SourceIP $remoteEP
        }

        # In Receive connector logs Exchange is the server:
        # event "<" = client (sender) sent a command to Exchange
        if ($event -ne '<') { continue }

        if ($data -like 'MAIL From:*' -or $data -like 'MAIL FROM:*') {
            $addr = Get-EmailAddress -Data $data -Command 'MAIL From'
            if ($null -ne $addr) {
                $s.MailFrom = $addr
                $s.HasMail  = $true
            }
        }
        elseif ($data -like 'RCPT To:*' -or $data -like 'RCPT TO:*') {
            $addr = Get-EmailAddress -Data $data -Command 'RCPT To'
            if ($null -ne $addr -and $s.RcptTo -notcontains $addr) {
                $s.RcptTo.Add($addr)
            }
        }
    }
}

Write-Progress -Activity "Processing log files" -Completed

#endregion

#region ── build result objects ─────────────────────────────────────────────────

$mailSessions = $sessions.Values |
    Where-Object { $_.HasMail -and $null -ne $_.MailFrom -and $_.RcptTo.Count -gt 0 } |
    Where-Object { $ExcludeIP.Count -eq 0 -or $_.SourceIP -notin $ExcludeIP }

$results = $mailSessions | ForEach-Object {
    foreach ($rcpt in $_.RcptTo) {
        [PSCustomObject]@{
            'Time (UTC)' = $_.FirstSeen
            'Server'     = $_.ServerName
            'Connector'  = $_.ConnectorName
            'Source IP'  = $_.SourceIP
            'From'       = $_.MailFrom
            'To'         = $rcpt
        }
    }
} | Sort-Object 'Time (UTC)'

#endregion

#region ── console output ───────────────────────────────────────────────────────

if ($results.Count -eq 0) {
    Write-Host ("No mail traffic found in the last {0}." -f $windowDesc) -ForegroundColor Yellow
}
else {
    Write-Host ("Found {0} message(s) across {1} session(s) in the last {2}:" -f
        $results.Count, ($mailSessions | Measure-Object).Count, $windowDesc) -ForegroundColor Green
    Write-Host ""
    $results | Format-Table -AutoSize -Wrap
}

#endregion

#region ── executive summary ────────────────────────────────────────────────────

$summaryLines = [System.Collections.Generic.List[string]]::new()

$summaryLines.Add('=' * 60)
$summaryLines.Add('EXECUTIVE SUMMARY')
$summaryLines.Add('=' * 60)
$summaryLines.Add("Period    : last $windowDesc")
$summaryLines.Add("From (UTC): $($cutoff.ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))")
$summaryLines.Add("To (UTC)  : $((Get-Date).ToUniversalTime().ToString('yyyy-MM-dd HH:mm:ss'))")
$summaryLines.Add("Log path  : $LogPath")
if ($Connector)             { $summaryLines.Add("Connector filter : $Connector") }
if ($ExcludeIP.Count -gt 0) { $summaryLines.Add("Excluded IPs     : $($ExcludeIP -join ', ')") }
$summaryLines.Add("")

if ($results.Count -eq 0) {
    $summaryLines.Add("No mail traffic found.")
}
else {
    $summaryLines.Add("Total messages : $($results.Count)")
    $summaryLines.Add("Total sessions : $(($mailSessions | Measure-Object).Count)")
    $summaryLines.Add("")

    # Group: Server -> Connector -> Source IP
    $byServer = $results | Group-Object 'Server' | Sort-Object Name

    foreach ($srvGroup in $byServer) {
        $summaryLines.Add("Server: $($srvGroup.Name)")
        $summaryLines.Add("-" * 56)

        $byConnector = $srvGroup.Group | Group-Object 'Connector' | Sort-Object Name

        foreach ($connGroup in $byConnector) {
            $summaryLines.Add("  Connector: $($connGroup.Name)")

            $byIP = $connGroup.Group | Group-Object 'Source IP' |
                Sort-Object { [int]$_.Count } -Descending

            foreach ($ipGroup in $byIP) {
                $summaryLines.Add(("    {0,-40} {1,5} message(s)" -f $ipGroup.Name, $ipGroup.Count))
            }
            $summaryLines.Add("")
        }
    }
}

$summaryLines.Add('=' * 60)

# Print to console
Write-Host ""
Write-Host ($summaryLines -join "`n") -ForegroundColor White

#endregion

#region ── file export ──────────────────────────────────────────────────────────

if ($ExportCsv) {
    try {
        $exportDir = Split-Path -Path $ExportCsv -Parent
        if ($exportDir -and -not (Test-Path $exportDir)) {
            New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
        }

        # Detail CSV
        if ($results.Count -gt 0) {
            $results | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Force
        }
        else {
            # Write an empty CSV with headers only
            [PSCustomObject]@{
                'Time (UTC)' = ''
                'Server'     = ''
                'Connector'  = ''
                'Source IP'  = ''
                'From'       = ''
                'To'         = ''
            } | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Force
        }

        # Summary TXT – same folder, same base name with _summary.txt suffix
        $baseName    = [System.IO.Path]::GetFileNameWithoutExtension($ExportCsv)
        $summaryPath = Join-Path $exportDir ($baseName + '_summary.txt')
        $summaryLines | Set-Content -Path $summaryPath -Encoding UTF8 -Force

        Write-Host ""
        Write-Host ("Detail CSV exported to : {0}" -f $ExportCsv)  -ForegroundColor Green
        Write-Host ("Summary exported to    : {0}" -f $summaryPath) -ForegroundColor Green
    }
    catch {
        Write-Error "Export failed: $_"
    }
}

#endregion
