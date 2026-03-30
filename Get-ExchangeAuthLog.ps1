<#
.SYNOPSIS
    Analyzes Exchange Server IMAP4 and POP3 protocol logs for authentication activity.

.DESCRIPTION
    Parses Exchange IMAP4 and POP3 protocol log files (W3C format) and reports:
      - Which accounts are authenticating
      - From which IP addresses
      - Via which protocol (IMAP4 / POP3)
      - Errors for failed authentication attempts

    Compatible with Exchange Server 2013, 2016 and 2019.

.PARAMETER ImapLogPath
    Path to the IMAP4 protocol log directory.
    Default: auto-detected from registry (%ExchangeInstallPath%\Logging\Imap4).

.PARAMETER PopLogPath
    Path to the POP3 protocol log directory.
    Default: auto-detected from registry (%ExchangeInstallPath%\Logging\Pop3).

.PARAMETER StartDate
    Analyze log files from this date/time onward. Default: 7 days ago.

.PARAMETER EndDate
    Analyze log files up to this date/time. Default: now.

.PARAMETER UserFilter
    Filter results by username (supports wildcards, e.g. "john*" or "*@contoso.com").

.PARAMETER IPFilter
    Filter results by client IP address (supports wildcards, e.g. "192.168.1.*").

.PARAMETER FailedOnly
    Show only failed authentication attempts.

.PARAMETER SuccessOnly
    Show only successful authentication attempts.

.PARAMETER ExportCsv
    Full path to export results as a CSV file.

.PARAMETER ExportHtml
    Full path to export results as an HTML report.

.PARAMETER Summary
    Show a summary table grouped by User / IP / Protocol instead of individual events.

.EXAMPLE
    .\Get-ExchangeAuthLog.ps1
    Analyze the last 7 days of IMAP4 and POP3 logs and display all events.

.EXAMPLE
    .\Get-ExchangeAuthLog.ps1 -StartDate (Get-Date).AddDays(-1) -FailedOnly
    Show only failed auth attempts from the last 24 hours.

.EXAMPLE
    .\Get-ExchangeAuthLog.ps1 -UserFilter "*@contoso.com" -ExportCsv C:\Reports\auth.csv
    Analyze all accounts at contoso.com and export to CSV.

.EXAMPLE
    .\Get-ExchangeAuthLog.ps1 -IPFilter "10.0.0.*" -Summary
    Show a summary of activity from a specific subnet.

.NOTES
    Requires Read access to the Exchange protocol log directories.
    Run on the Exchange server, or map the log paths via UNC.
#>

[CmdletBinding()]
param(
    [string]$ImapLogPath,
    [string]$PopLogPath,
    [datetime]$StartDate = (Get-Date).AddDays(-7),
    [datetime]$EndDate   = (Get-Date),
    [string]$UserFilter,
    [string]$IPFilter,
    [switch]$FailedOnly,
    [switch]$SuccessOnly,
    [string]$ExportCsv,
    [string]$ExportHtml,
    [switch]$Summary
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region ── Helpers ────────────────────────────────────────────────────────────

function Resolve-ExchangeLogPath {
    <#
    Tries to locate the Exchange installation path via the registry.
    Returns $null if Exchange is not found.
    #>
    try {
        $regPath = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
        $installPath = (Get-ItemProperty $regPath -ErrorAction Stop).MsiInstallPath
        return $installPath
    }
    catch {
        return $null
    }
}

function Parse-W3CLogFile {
    <#
    Reads a W3C Extended Log Format file and returns an array of hashtables,
    each representing one log record keyed by the field names from #Fields:.
    Comment lines (#Software, #Version, #Date) are skipped.
    The #Fields: line is re-read whenever it changes (log rotation).
    #>
    param(
        [string]$Path
    )

    $fieldNames = @()
    $records    = [System.Collections.Generic.List[hashtable]]::new()

    foreach ($line in [System.IO.File]::ReadLines($Path)) {
        if ($line.StartsWith('#Fields:')) {
            # Parse header: "#Fields: field1,field2,..."
            $fieldNames = $line.Substring(8).Trim() -split ','
            continue
        }
        if ($line.StartsWith('#') -or [string]::IsNullOrWhiteSpace($line)) {
            continue
        }
        if ($fieldNames.Count -eq 0) { continue }

        # Exchange logs are comma-delimited; quoted fields can contain commas.
        # Use a simple CSV split that respects double-quotes.
        $values = Split-CsvLine $line

        $record = @{}
        for ($i = 0; $i -lt $fieldNames.Count; $i++) {
            $record[$fieldNames[$i]] = if ($i -lt $values.Count) { $values[$i] } else { '' }
        }
        $records.Add($record)
    }
    return $records
}

function Split-CsvLine {
    <#
    Minimal CSV line splitter that handles double-quoted fields.
    #>
    param([string]$Line)

    $result  = [System.Collections.Generic.List[string]]::new()
    $current = [System.Text.StringBuilder]::new()
    $inQuote = $false

    for ($i = 0; $i -lt $Line.Length; $i++) {
        $ch = $Line[$i]
        if ($ch -eq '"') {
            if ($inQuote -and $i + 1 -lt $Line.Length -and $Line[$i + 1] -eq '"') {
                # Escaped quote ""
                [void]$current.Append('"')
                $i++
            }
            else {
                $inQuote = -not $inQuote
            }
        }
        elseif ($ch -eq ',' -and -not $inQuote) {
            $result.Add($current.ToString())
            [void]$current.Clear()
        }
        else {
            [void]$current.Append($ch)
        }
    }
    $result.Add($current.ToString())
    return $result
}

function Extract-IPFromEndpoint {
    <#
    Exchange logs endpoints as "1.2.3.4:143" or "[::1]:143".
    Returns only the IP part.
    #>
    param([string]$Endpoint)

    if ([string]::IsNullOrEmpty($Endpoint) -or $Endpoint -eq '-') { return '-' }

    # IPv6 in brackets
    if ($Endpoint -match '^\[(.+)\]:\d+$') { return $Matches[1] }

    # IPv4 with port
    if ($Endpoint -match '^([\d.]+):\d+$')  { return $Matches[1] }

    return $Endpoint
}

#endregion

#region ── IMAP4 Log Parser ───────────────────────────────────────────────────

function Parse-ImapLogs {
    param(
        [string]$LogDirectory,
        [datetime]$Start,
        [datetime]$End
    )

    $authEvents = [System.Collections.Generic.List[pscustomobject]]::new()

    if (-not (Test-Path $LogDirectory)) {
        Write-Warning "IMAP4 log directory not found: $LogDirectory"
        return $authEvents
    }

    $logFiles = Get-ChildItem -Path $LogDirectory -Filter '*.log' -File |
                Where-Object { $_.LastWriteTime -ge $Start.AddDays(-1) -and $_.LastWriteTime -le $End.AddDays(1) }

    if (-not $logFiles) {
        Write-Verbose "No IMAP4 log files found in range."
        return $authEvents
    }

    Write-Host "  Parsing $($logFiles.Count) IMAP4 log file(s)..." -ForegroundColor Cyan

    foreach ($file in $logFiles) {
        Write-Verbose "  Reading: $($file.Name)"

        # Per-session state: sessionId -> { ClientIP, User, LastLoginCmd, LastAuthMethod }
        $sessions = @{}

        try {
            $records = Parse-W3CLogFile -Path $file.FullName
        }
        catch {
            Write-Warning "Failed to read $($file.Name): $_"
            continue
        }

        foreach ($rec in $records) {
            # Timestamp field is usually "date-time" or "DateTime"
            $tsField = if ($rec.ContainsKey('date-time')) { 'date-time' } else { 'DateTime' }
            if (-not $rec.ContainsKey($tsField)) { continue }

            $ts = [datetime]::MinValue
            if (-not [datetime]::TryParse($rec[$tsField], [ref]$ts)) { continue }
            if ($ts -lt $Start -or $ts -gt $End) { continue }

            $sessionId = if ($rec.ContainsKey('session')) { $rec['session'] } else { '' }
            $event     = if ($rec.ContainsKey('event'))   { $rec['event'].ToUpper() } else { '' }
            $data      = if ($rec.ContainsKey('data'))    { $rec['data'] }  else { '' }
            $remoteEp  = if ($rec.ContainsKey('remote-endpoint')) { $rec['remote-endpoint'] } else { '-' }
            $clientIP  = Extract-IPFromEndpoint $remoteEp

            # Initialise session tracking
            if ($sessionId -and -not $sessions.ContainsKey($sessionId)) {
                $sessions[$sessionId] = @{
                    ClientIP       = $clientIP
                    User           = ''
                    PendingLogin   = $false
                    AuthMethod     = ''
                    LastCmdTag     = ''
                }
            }
            $sess = if ($sessionId) { $sessions[$sessionId] } else { @{ ClientIP = $clientIP; User = ''; PendingLogin = $false; AuthMethod = ''; LastCmdTag = '' } }

            # Update IP (might change in proxy scenarios)
            if ($clientIP -ne '-') { $sess.ClientIP = $clientIP }

            switch ($event) {
                'COMMAND' {
                    # IMAP command line, e.g.:  "A001 LOGIN user@domain *"
                    #                     or:  "A001 AUTHENTICATE PLAIN"
                    if ($data -match '^(\S+)\s+(LOGIN|AUTHENTICATE)\s+(\S+)') {
                        $sess.LastCmdTag   = $Matches[1]
                        $sess.AuthMethod   = $Matches[2].ToUpper()
                        $rawUser           = $Matches[3]
                        # Password is 4th token for LOGIN; strip it
                        $sess.User         = $rawUser -replace '\s+\S+$', ''
                        $sess.PendingLogin  = $true
                    }
                }
                'RESPONSE' {
                    if ($sess.PendingLogin) {
                        # Tagged OK = success, NO or BAD = failure
                        $success = $false
                        $errMsg  = ''

                        if ($data -match "^$([regex]::Escape($sess.LastCmdTag))\s+OK\b") {
                            $success = $true
                        }
                        elseif ($data -match "^$([regex]::Escape($sess.LastCmdTag))\s+(NO|BAD)\b(.*)") {
                            $errMsg = $Matches[2].Trim() -replace '^\[.*?\]\s*', ''
                            if (-not $errMsg) { $errMsg = $data }
                        }
                        else {
                            # Untagged continuation – skip
                            continue
                        }

                        $authEvents.Add([pscustomobject]@{
                            Timestamp  = $ts
                            Protocol   = 'IMAP4'
                            ClientIP   = $sess.ClientIP
                            Username   = $sess.User
                            AuthMethod = $sess.AuthMethod
                            Success    = $success
                            Error      = if ($success) { '' } else { if ($errMsg) { $errMsg } else { 'Authentication failed' } }
                            LogFile    = $file.Name
                        })

                        $sess.PendingLogin = $false
                        $sess.User         = ''
                        $sess.AuthMethod   = ''
                        $sess.LastCmdTag   = ''
                    }
                }
                'CONNECT' {
                    # Refresh IP on new connection
                    if ($clientIP -ne '-') { $sess.ClientIP = $clientIP }
                }
            }
        }
    }

    return $authEvents
}

#endregion

#region ── POP3 Log Parser ────────────────────────────────────────────────────

function Parse-Pop3Logs {
    param(
        [string]$LogDirectory,
        [datetime]$Start,
        [datetime]$End
    )

    $authEvents = [System.Collections.Generic.List[pscustomobject]]::new()

    if (-not (Test-Path $LogDirectory)) {
        Write-Warning "POP3 log directory not found: $LogDirectory"
        return $authEvents
    }

    $logFiles = Get-ChildItem -Path $LogDirectory -Filter '*.log' -File |
                Where-Object { $_.LastWriteTime -ge $Start.AddDays(-1) -and $_.LastWriteTime -le $End.AddDays(1) }

    if (-not $logFiles) {
        Write-Verbose "No POP3 log files found in range."
        return $authEvents
    }

    Write-Host "  Parsing $($logFiles.Count) POP3 log file(s)..." -ForegroundColor Cyan

    foreach ($file in $logFiles) {
        Write-Verbose "  Reading: $($file.Name)"

        $sessions = @{}

        try {
            $records = Parse-W3CLogFile -Path $file.FullName
        }
        catch {
            Write-Warning "Failed to read $($file.Name): $_"
            continue
        }

        foreach ($rec in $records) {
            $tsField = if ($rec.ContainsKey('date-time')) { 'date-time' } else { 'DateTime' }
            if (-not $rec.ContainsKey($tsField)) { continue }

            $ts = [datetime]::MinValue
            if (-not [datetime]::TryParse($rec[$tsField], [ref]$ts)) { continue }
            if ($ts -lt $Start -or $ts -gt $End) { continue }

            $sessionId = if ($rec.ContainsKey('session')) { $rec['session'] } else { '' }
            $event     = if ($rec.ContainsKey('event'))   { $rec['event'].ToUpper() } else { '' }
            $data      = if ($rec.ContainsKey('data'))    { $rec['data'] }  else { '' }
            $remoteEp  = if ($rec.ContainsKey('remote-endpoint')) { $rec['remote-endpoint'] } else { '-' }
            $clientIP  = Extract-IPFromEndpoint $remoteEp

            if ($sessionId -and -not $sessions.ContainsKey($sessionId)) {
                $sessions[$sessionId] = @{
                    ClientIP      = $clientIP
                    User          = ''
                    PassSent      = $false
                    AuthTimestamp = $ts
                }
            }
            $sess = if ($sessionId) { $sessions[$sessionId] } else { @{ ClientIP = $clientIP; User = ''; PassSent = $false; AuthTimestamp = $ts } }

            if ($clientIP -ne '-') { $sess.ClientIP = $clientIP }

            switch ($event) {
                'COMMAND' {
                    # POP3 command lines: "USER user@domain" or "PASS *" or "AUTH PLAIN"
                    if ($data -match '^USER\s+(\S+)') {
                        $sess.User          = $Matches[1]
                        $sess.PassSent      = $false
                        $sess.AuthTimestamp = $ts
                    }
                    elseif ($data -match '^PASS\b') {
                        $sess.PassSent = $true
                    }
                    elseif ($data -match '^AUTH\s+(\S+)') {
                        # SASL AUTH (AUTH PLAIN, AUTH LOGIN, etc.)
                        $sess.PassSent      = $true   # AUTH is a single round-trip for POP3
                        $sess.AuthTimestamp = $ts
                    }
                }
                'RESPONSE' {
                    if ($sess.PassSent -and $sess.User) {
                        $success = $false
                        $errMsg  = ''

                        if ($data -match '^\+OK\b') {
                            $success = $true
                        }
                        elseif ($data -match '^-ERR\s*(.*)') {
                            $errMsg = $Matches[1].Trim()
                        }
                        else {
                            continue
                        }

                        $authEvents.Add([pscustomobject]@{
                            Timestamp  = $sess.AuthTimestamp
                            Protocol   = 'POP3'
                            ClientIP   = $sess.ClientIP
                            Username   = $sess.User
                            AuthMethod = 'USER/PASS'
                            Success    = $success
                            Error      = if ($success) { '' } else { if ($errMsg) { $errMsg } else { 'Authentication failed' } }
                            LogFile    = $file.Name
                        })

                        $sess.PassSent = $false
                        $sess.User     = ''
                    }
                }
                'CONNECT' {
                    if ($clientIP -ne '-') { $sess.ClientIP = $clientIP }
                }
            }
        }
    }

    return $authEvents
}

#endregion

#region ── HTML Report ────────────────────────────────────────────────────────

function Export-HtmlReport {
    param(
        [pscustomobject[]]$Events,
        [string]$Path,
        [datetime]$Start,
        [datetime]$End
    )

    $successCount = ($Events | Where-Object Success).Count
    $failCount    = ($Events | Where-Object { -not $_.Success }).Count
    $userCount    = ($Events | Select-Object -ExpandProperty Username -Unique).Count
    $ipCount      = ($Events | Select-Object -ExpandProperty ClientIP -Unique).Count

    $rows = foreach ($e in $Events) {
        $statusStyle = if ($e.Success) { 'background:#d4edda;color:#155724' } else { 'background:#f8d7da;color:#721c24' }
        $statusText  = if ($e.Success) { 'SUCCESS' } else { 'FAILED' }
        $errorCell   = if ($e.Error)   { "<span style='color:#c0392b'>$([System.Web.HttpUtility]::HtmlEncode($e.Error))</span>" } else { '' }
        "
        <tr>
          <td>$($e.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))</td>
          <td><strong>$($e.Protocol)</strong></td>
          <td>$($e.ClientIP)</td>
          <td>$([System.Web.HttpUtility]::HtmlEncode($e.Username))</td>
          <td>$($e.AuthMethod)</td>
          <td style='$statusStyle;font-weight:bold;text-align:center'>$statusText</td>
          <td>$errorCell</td>
        </tr>"
    }

    $html = @"
<!DOCTYPE html>
<html lang='en'>
<head>
<meta charset='UTF-8'>
<title>Exchange Auth Log Report</title>
<style>
  body { font-family: Segoe UI, Arial, sans-serif; margin: 20px; background: #f5f5f5; }
  h1   { color: #0078d4; }
  .summary { display:flex; gap:20px; margin-bottom:20px; }
  .card    { background:#fff; border-radius:8px; padding:15px 25px; box-shadow:0 2px 4px rgba(0,0,0,.1); text-align:center; min-width:120px; }
  .card .num  { font-size:2em; font-weight:bold; }
  .card .lbl  { color:#555; font-size:.85em; }
  table  { width:100%; border-collapse:collapse; background:#fff; border-radius:8px; overflow:hidden; box-shadow:0 2px 4px rgba(0,0,0,.1); }
  th     { background:#0078d4; color:#fff; padding:10px 12px; text-align:left; font-size:.9em; }
  td     { padding:8px 12px; border-bottom:1px solid #eee; font-size:.88em; }
  tr:last-child td { border-bottom:none; }
  tr:hover td { background:#f0f8ff; }
  .period { color:#555; font-size:.9em; margin-bottom:15px; }
</style>
</head>
<body>
<h1>Exchange IMAP4 / POP3 Authentication Report</h1>
<p class='period'>Period: $($Start.ToString('yyyy-MM-dd HH:mm')) &ndash; $($End.ToString('yyyy-MM-dd HH:mm'))</p>
<div class='summary'>
  <div class='card'><div class='num'>$($Events.Count)</div><div class='lbl'>Total Events</div></div>
  <div class='card' style='border-top:4px solid #28a745'><div class='num' style='color:#28a745'>$successCount</div><div class='lbl'>Successful</div></div>
  <div class='card' style='border-top:4px solid #dc3545'><div class='num' style='color:#dc3545'>$failCount</div><div class='lbl'>Failed</div></div>
  <div class='card'><div class='num'>$userCount</div><div class='lbl'>Unique Users</div></div>
  <div class='card'><div class='num'>$ipCount</div><div class='lbl'>Unique IPs</div></div>
</div>
<table>
  <thead>
    <tr>
      <th>Timestamp</th><th>Protocol</th><th>Client IP</th><th>Username</th><th>Auth Method</th><th>Status</th><th>Error</th>
    </tr>
  </thead>
  <tbody>
    $($rows -join "`n")
  </tbody>
</table>
<p style='color:#999;font-size:.8em;margin-top:15px'>Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
</body>
</html>
"@

    # HtmlEncode needs System.Web – available in .NET Framework; on .NET Core use manual encode
    Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

    $html | Set-Content -Path $Path -Encoding UTF8
    Write-Host "  HTML report saved: $Path" -ForegroundColor Green
}

#endregion

#region ── Main ───────────────────────────────────────────────────────────────

Write-Host ""
Write-Host "================================================================" -ForegroundColor DarkCyan
Write-Host "  Exchange IMAP4 / POP3 Authentication Log Analyzer" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor DarkCyan
Write-Host "  Period : $($StartDate.ToString('yyyy-MM-dd HH:mm')) - $($EndDate.ToString('yyyy-MM-dd HH:mm'))"
Write-Host ""

# ── Resolve log paths ────────────────────────────────────────────────────────
$exchangeRoot = Resolve-ExchangeLogPath

if (-not $ImapLogPath) {
    if ($exchangeRoot) {
        $ImapLogPath = Join-Path $exchangeRoot 'Logging\Imap4'
    }
    else {
        $ImapLogPath = 'C:\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4'
        Write-Warning "Exchange install path not found in registry. Trying default: $ImapLogPath"
    }
}

if (-not $PopLogPath) {
    if ($exchangeRoot) {
        $PopLogPath = Join-Path $exchangeRoot 'Logging\Pop3'
    }
    else {
        $PopLogPath = 'C:\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3'
        Write-Warning "Exchange install path not found in registry. Trying default: $PopLogPath"
    }
}

Write-Host "  IMAP4 logs : $ImapLogPath"
Write-Host "  POP3 logs  : $PopLogPath"
Write-Host ""

# ── Parse logs ───────────────────────────────────────────────────────────────
$allEvents = [System.Collections.Generic.List[pscustomobject]]::new()

Write-Host "[1/2] Processing IMAP4 logs..." -ForegroundColor Yellow
$imapEvents = Parse-ImapLogs -LogDirectory $ImapLogPath -Start $StartDate -End $EndDate
$allEvents.AddRange($imapEvents)

Write-Host "[2/2] Processing POP3 logs..." -ForegroundColor Yellow
$popEvents = Parse-Pop3Logs -LogDirectory $PopLogPath -Start $StartDate -End $EndDate
$allEvents.AddRange($popEvents)

Write-Host ""

if ($allEvents.Count -eq 0) {
    Write-Warning "No authentication events found in the specified time range."
    exit 0
}

# ── Apply filters ─────────────────────────────────────────────────────────────
$filtered = $allEvents

if ($UserFilter) {
    $filtered = $filtered | Where-Object { $_.Username -like $UserFilter }
}
if ($IPFilter) {
    $filtered = $filtered | Where-Object { $_.ClientIP -like $IPFilter }
}
if ($FailedOnly)  { $filtered = $filtered | Where-Object { -not $_.Success } }
if ($SuccessOnly) { $filtered = $filtered | Where-Object { $_.Success } }

$filtered = @($filtered | Sort-Object Timestamp)

if ($filtered.Count -eq 0) {
    Write-Warning "No events match the specified filters."
    exit 0
}

# ── Display results ───────────────────────────────────────────────────────────
$successCount = ($filtered | Where-Object { $_.Success }).Count
$failCount    = ($filtered | Where-Object { -not $_.Success }).Count

Write-Host "================================================================" -ForegroundColor DarkCyan
Write-Host "  Results: $($filtered.Count) events  |  Success: $successCount  |  Failed: $failCount" -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor DarkCyan
Write-Host ""

if ($Summary) {
    # ── Summary view: group by User / IP / Protocol ──────────────────────────
    Write-Host "SUMMARY (grouped by Username / IP / Protocol)" -ForegroundColor Yellow
    Write-Host ""

    $grouped = $filtered |
        Group-Object Username, ClientIP, Protocol |
        ForEach-Object {
            $grpEvents  = $_.Group
            $ok         = ($grpEvents | Where-Object { $_.Success }).Count
            $fail       = ($grpEvents | Where-Object { -not $_.Success }).Count
            $firstSeen  = ($grpEvents | Sort-Object Timestamp | Select-Object -First 1).Timestamp
            $lastSeen   = ($grpEvents | Sort-Object Timestamp | Select-Object -Last  1).Timestamp
            $topErrors  = ($grpEvents | Where-Object { $_.Error } | Group-Object Error | Sort-Object Count -Descending | Select-Object -First 3 | ForEach-Object { "$($_.Name) ($($_.Count)x)" }) -join '; '

            [pscustomobject]@{
                Username   = $grpEvents[0].Username
                ClientIP   = $grpEvents[0].ClientIP
                Protocol   = $grpEvents[0].Protocol
                Success    = $ok
                Failed     = $fail
                Total      = $grpEvents.Count
                FirstSeen  = $firstSeen.ToString('yyyy-MM-dd HH:mm')
                LastSeen   = $lastSeen.ToString('yyyy-MM-dd HH:mm')
                TopErrors  = $topErrors
            }
        } |
        Sort-Object Failed -Descending

    $grouped | Format-Table -AutoSize

    # Highlight accounts with high failure counts
    $suspicious = $grouped | Where-Object { $_.Failed -ge 5 }
    if ($suspicious) {
        Write-Host "WARNING: Accounts with 5+ failed auth attempts:" -ForegroundColor Red
        $suspicious | ForEach-Object {
            Write-Host "  $($_.Username) from $($_.ClientIP) via $($_.Protocol) - $($_.Failed) failures" -ForegroundColor Red
        }
        Write-Host ""
    }
}
else {
    # ── Detailed event view ───────────────────────────────────────────────────
    foreach ($event in $filtered) {
        $statusColor = if ($event.Success) { 'Green' } else { 'Red' }
        $statusText  = if ($event.Success) { '[  OK  ]' } else { '[FAILED]' }

        Write-Host "$statusText " -NoNewline -ForegroundColor $statusColor
        Write-Host "$($event.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))  " -NoNewline -ForegroundColor Gray
        Write-Host "$($event.Protocol.PadRight(5))  " -NoNewline -ForegroundColor Cyan
        Write-Host "$($event.ClientIP.PadRight(16))  " -NoNewline
        Write-Host "$($event.Username)" -NoNewline -ForegroundColor Yellow

        if (-not $event.Success -and $event.Error) {
            Write-Host "  --> $($event.Error)" -ForegroundColor DarkRed
        }
        else {
            Write-Host ""
        }
    }
    Write-Host ""
}

# ── Top offenders table (always show if there are failures) ──────────────────
$failedEvents = $filtered | Where-Object { -not $_.Success }
if ($failedEvents -and -not $FailedOnly -and -not $Summary) {
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkYellow
    Write-Host "TOP FAILED AUTHENTICATION ATTEMPTS" -ForegroundColor Yellow
    Write-Host "----------------------------------------------------------------" -ForegroundColor DarkYellow

    $failedEvents |
        Group-Object ClientIP, Username |
        ForEach-Object {
            $topErr = ($_.Group | Group-Object Error | Sort-Object Count -Descending | Select-Object -First 1).Name
            [pscustomobject]@{
                ClientIP  = $_.Group[0].ClientIP
                Username  = $_.Group[0].Username
                Protocol  = ($_.Group | Select-Object -ExpandProperty Protocol -Unique) -join '/'
                Count     = $_.Count
                LastSeen  = ($_.Group | Sort-Object Timestamp | Select-Object -Last 1).Timestamp.ToString('yyyy-MM-dd HH:mm:ss')
                TopError  = $topErr
            }
        } |
        Sort-Object Count -Descending |
        Select-Object -First 20 |
        Format-Table -AutoSize
}

# ── Export ────────────────────────────────────────────────────────────────────
if ($ExportCsv) {
    $filtered | Export-Csv -Path $ExportCsv -NoTypeInformation -Encoding UTF8
    Write-Host "CSV exported: $ExportCsv" -ForegroundColor Green
}

if ($ExportHtml) {
    Export-HtmlReport -Events $filtered -Path $ExportHtml -Start $StartDate -End $EndDate
}

Write-Host ""
Write-Host "Done. Total events analyzed: $($allEvents.Count)" -ForegroundColor DarkCyan

#endregion
