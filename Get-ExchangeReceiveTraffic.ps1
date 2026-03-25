#Requires -Version 5.0

<#
.SYNOPSIS
    Analyzes Exchange SMTP protocol logs (Receive or Send connectors) for mail traffic.

.DESCRIPTION
    Parses Exchange SMTP protocol log files (.log) and reports mail traffic per
    connector, including sender address, recipient(s), remote IP, and connector name.

    Supports both Receive connector logs (Exchange acts as server) and Send connector
    logs (Exchange acts as client). Log type is auto-detected from the #Log-type header.

    A new log file is created every hour by Exchange. Files are selected based on
    their last-write time to cover the requested time window.

    An executive summary (per server -> per IP with message counts) is always shown
    on the console. When -ExportCsv is used, the summary is also saved as a .txt file
    with the same base name.

.PARAMETER LogPath
    REQUIRED. Path to the folder containing Exchange SMTP protocol log files.
    Examples:
      Receive: C:\...\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpReceive
      Send   : C:\...\TransportRoles\Logs\FrontEnd\ProtocolLog\SmtpSend

.PARAMETER Hours
    Number of hours back from now to analyse. Cannot be combined with -Days.
    If neither -Hours nor -Days is specified, defaults to 5 hours.

.PARAMETER Days
    Number of days back from now to analyse. Cannot be combined with -Hours.

.PARAMETER Connector
    Optional. Filter results to a specific connector name (partial, case-insensitive match).

.PARAMETER ExportCsv
    Optional. Full path to a CSV file for detailed results.
    A matching summary text file (<basename>_summary.txt) is created automatically.
    Both files are overwritten if they already exist.

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -LogPath "C:\Logs\SmtpReceive"

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -LogPath "C:\Logs\SmtpReceive" -Hours 12

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -LogPath "C:\Logs\SmtpSend" -Days 3

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -LogPath "C:\Logs\SmtpReceive" -Hours 5 -Connector "Anon Relay"

.EXAMPLE
    .\Get-ExchangeReceiveTraffic.ps1 -LogPath "C:\Logs\SmtpReceive" -Days 7 -ExportCsv "C:\Reports\traffic.csv"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$LogPath,

    [Parameter()]
    [ValidateRange(1, 8760)]
    [int]$Hours = 0,

    [Parameter()]
    [ValidateRange(1, 365)]
    [int]$Days = 0,

    [Parameter()]
    [string]$Connector = '',

    [Parameter()]
    [string]$ExportCsv = ''
)

#region ── validation ───────────────────────────────────────────────────────────

if ($Hours -gt 0 -and $Days -gt 0) {
    Write-Error "Specify either -Hours or -Days, not both."
    exit 1
}

if ($Days -gt 0) {
    $cutoff      = (Get-Date).AddDays(-$Days)
    $windowDesc  = "$Days day(s)"
}
elseif ($Hours -gt 0) {
    $cutoff      = (Get-Date).AddHours(-$Hours)
    $windowDesc  = "$Hours hour(s)"
}
else {
    $cutoff      = (Get-Date).AddHours(-5)
    $windowDesc  = '5 hours (default)'
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

function Get-RemoteIP {
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

Write-Host "Exchange SMTP Connector Traffic Analyzer" -ForegroundColor Cyan
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

# Holds detected log type (populated from first file that has the header)
$detectedLogType = $null   # 'Receive' or 'Send'

# Key: session-id  Value: hashtable with session data
$sessions = @{}

foreach ($file in $logFiles) {
    Write-Verbose "Processing: $($file.Name)"

    $lines = Get-Content -Path $file.FullName -Encoding UTF8 -ErrorAction SilentlyContinue
    if (-not $lines) { continue }

    # Detect log type from this file (use first detection found)
    if (-not $detectedLogType) {
        $logTypeLine = $lines | Where-Object { $_ -like '#Log-type:*' } | Select-Object -First 1
        if ($logTypeLine -match 'Send') {
            $detectedLogType = 'Send'
        }
        elseif ($logTypeLine -match 'Receive') {
            $detectedLogType = 'Receive'
        }
    }

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
                RemoteIP      = Get-RemoteIP      $remoteEP
                FirstSeen     = $timestamp
                MailFrom      = $null
                RcptTo        = [System.Collections.Generic.List[string]]::new()
                HasMail       = $false
            }
        }

        $s = $sessions[$sessionId]

        # Capture remote IP from the first non-empty endpoint seen for this session
        if (-not $s.RemoteIP -and $remoteEP) {
            $s.RemoteIP = Get-RemoteIP $remoteEP
        }

        # Event direction depends on log type:
        #   Receive log – Exchange is server  → client commands arrive as event "<"
        #   Send log    – Exchange is client  → Exchange commands go out as event ">"
        $cmdEvent = if ($detectedLogType -eq 'Send') { '>' } else { '<' }
        if ($event -ne $cmdEvent) { continue }

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

if (-not $detectedLogType) { $detectedLogType = 'Receive' }   # fallback

# Label the IP column based on log type
$ipColumnLabel = if ($detectedLogType -eq 'Send') { 'Destination IP' } else { 'Source IP' }

Write-Host ("Log type detected  : SMTP {0} Protocol Log" -f $detectedLogType) -ForegroundColor Cyan
Write-Host ""

#endregion

#region ── build result objects ─────────────────────────────────────────────────

$mailSessions = $sessions.Values |
    Where-Object { $_.HasMail -and $null -ne $_.MailFrom -and $_.RcptTo.Count -gt 0 }

$results = $mailSessions | ForEach-Object {
    foreach ($rcpt in $_.RcptTo) {
        [PSCustomObject]@{
            'Time (UTC)'   = $_.FirstSeen
            'Server'       = $_.ServerName
            'Connector'    = $_.ConnectorName
            $ipColumnLabel = $_.RemoteIP
            'From'         = $_.MailFrom
            'To'           = $rcpt
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
$summaryLines.Add("Log type  : SMTP $detectedLogType Protocol Log")
$summaryLines.Add("Log path  : $LogPath")
if ($Connector) { $summaryLines.Add("Connector filter: $Connector") }
$summaryLines.Add("")

if ($results.Count -eq 0) {
    $summaryLines.Add("No mail traffic found.")
}
else {
    $summaryLines.Add("Total messages : $($results.Count)")
    $summaryLines.Add("Total sessions : $(($mailSessions | Measure-Object).Count)")
    $summaryLines.Add("")

    # Group: Server → Connector → IP
    $byServer = $results | Group-Object 'Server' | Sort-Object Name

    foreach ($srvGroup in $byServer) {
        $summaryLines.Add("Server: $($srvGroup.Name)")
        $summaryLines.Add("-" * 56)

        $byConnector = $srvGroup.Group | Group-Object 'Connector' | Sort-Object Name

        foreach ($connGroup in $byConnector) {
            $summaryLines.Add("  Connector: $($connGroup.Name)")

            $byIP = $connGroup.Group | Group-Object $ipColumnLabel |
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
                'Time (UTC)'   = ''
                'Server'       = ''
                'Connector'    = ''
                $ipColumnLabel = ''
                'From'         = ''
                'To'           = ''
            } | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8 -Force
        }

        # Summary TXT – same folder, same base name with _summary.txt suffix
        $baseName    = [System.IO.Path]::GetFileNameWithoutExtension($ExportCsv)
        $summaryPath = Join-Path $exportDir ($baseName + '_summary.txt')
        $summaryLines | Set-Content -Path $summaryPath -Encoding UTF8 -Force

        Write-Host ""
        Write-Host ("Detail CSV exported to : {0}" -f $ExportCsv)       -ForegroundColor Green
        Write-Host ("Summary exported to    : {0}" -f $summaryPath)      -ForegroundColor Green
    }
    catch {
        Write-Error "Export failed: $_"
    }
}

#endregion
