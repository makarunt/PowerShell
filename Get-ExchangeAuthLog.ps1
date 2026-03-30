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

    $fieldNames  = @()
    $records     = [System.Collections.Generic.List[hashtable]]::new()
    $isFirstLine = $true

    # Open with FileShare.ReadWrite so we can read files Exchange currently has open for writing.
    $fs     = [System.IO.FileStream]::new($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
    $reader = [System.IO.StreamReader]::new($fs, [System.Text.Encoding]::UTF8)
    try {
    $line = $null
    while (($line = $reader.ReadLine()) -ne $null) {
        if ($line.StartsWith('#Fields:')) {
            # Standard W3C fields header: "#Fields: field1,field2,..."
            $fieldNames  = $line.Substring(8).Trim() -split ','
            $isFirstLine = $false
            continue
        }
        if ($line.StartsWith('#') -or [string]::IsNullOrWhiteSpace($line)) {
            $isFirstLine = $false
            continue
        }

        # Exchange logs start with an undecorated field-name line (no '#Fields:' prefix)
        # before the #Software/#Version/etc. comment block.  Detect it by checking that
        # we haven't seen any data yet and that every comma-separated token looks like
        # a valid identifier (letters/digits only – no spaces, no colons).
        if ($isFirstLine -and $fieldNames.Count -eq 0) {
            $isFirstLine = $false
            $tokens = $line -split ','
            $looksLikeHeader = $tokens.Count -ge 3 -and
                               (@($tokens | Where-Object { $_ -match '[^A-Za-z0-9_]' })).Count -eq 0
            if ($looksLikeHeader) {
                $fieldNames = $tokens
                continue
            }
        }
        $isFirstLine = $false

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
    } # end while
    finally {
        $reader.Dispose()
        $fs.Dispose()
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
    Exchange logs endpoints as "1.2.3.4:995" or "[::1]:993".
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

function Get-ContextResult {
    <#
    Exchange protocol logs store the result of each command in the 'context' field.
    Format: R=OK   or   R=OK;Msg="Proxy:...;ProxySuccess";ActivityContextData=...
            R=SomeErrorCode;Msg="Error description";...

    Returns a hashtable: @{ Success = $true/$false; Error = 'message or empty' }
    #>
    param([string]$Context)

    if ([string]::IsNullOrEmpty($Context)) {
        return @{ Success = $false; Error = 'No response context' }
    }

    # Extract the R= value (first token up to ; or end of string)
    $rValue = ''
    if ($Context -match '(?:^|,)R=([^;,]+)') {
        $rValue = $Matches[1].Trim()
    }

    if ($rValue -eq 'OK') {
        return @{ Success = $true; Error = '' }
    }

    # Extract Msg= value for a readable error description
    $msg = ''
    if ($Context -match 'Msg=""([^""]+)""') {
        $msg = $Matches[1].Trim()
        # Proxy success buried in R=OK scenarios shouldn't appear here, but clean up proxy strings
        $msg = $msg -replace ';ProxySuccess$', '' -replace '^Proxy:[^;]+;', ''
    }

    $error = if ($msg) { "$rValue - $msg" } elseif ($rValue) { $rValue } else { $Context }
    return @{ Success = $false; Error = $error }
}

function Get-RecordField {
    <#
    Safe field accessor for log record hashtables.
    Returns empty string if field doesn't exist.
    #>
    param([hashtable]$Record, [string[]]$Names)
    foreach ($n in $Names) {
        if ($Record.ContainsKey($n) -and $Record[$n] -ne '-') { return $Record[$n] }
    }
    return ''
}

#endregion

#region ── IMAP4 Log Parser ───────────────────────────────────────────────────

function Parse-ImapLogs {
    <#
    Exchange IMAP4 log format (Exchange 2013/2016/2019):
      Fields: dateTime, sessionId, seqNumber, sIp, cIp, user,
              duration, rqsize, rpsize, command, parameters, context, puid

    Authentication commands:
      login       - LOGIN username *  (password masked as *)
                    context contains result: R=OK or R=<ErrorCode>
      authenticate - AUTHENTICATE PLAIN/LOGIN/NTLM
                    user field contains authenticated username
                    context contains result

    Both frontend (IMAP4) and backend (IMAP4BE) logs share this format.
    #>
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

        try {
            $records = Parse-W3CLogFile -Path $file.FullName
        }
        catch {
            Write-Warning "Failed to read $($file.Name): $_"
            continue
        }

        foreach ($rec in $records) {
            # ── Timestamp ────────────────────────────────────────────────────
            $tsRaw = Get-RecordField $rec 'dateTime','date-time','DateTime'
            $ts    = [datetime]::MinValue
            if (-not $tsRaw -or -not [datetime]::TryParse($tsRaw, [ref]$ts)) { continue }
            if ($ts -lt $Start -or $ts -gt $End) { continue }

            # ── Fields ───────────────────────────────────────────────────────
            $cmd       = (Get-RecordField $rec 'command').ToLower()
            $params    = Get-RecordField $rec 'parameters'
            $context   = Get-RecordField $rec 'context'
            $userField = Get-RecordField $rec 'user'
            $clientIP  = Extract-IPFromEndpoint (Get-RecordField $rec 'cIp','remote-endpoint')

            # ── LOGIN command: "username *"  (password is masked) ────────────
            if ($cmd -eq 'login') {
                $username = if ($params -match '^(\S+)\s+\*') { $Matches[1] }
                            elseif ($params)                   { $params }
                            else                               { $userField }

                $result = Get-ContextResult $context
                $authEvents.Add([pscustomobject]@{
                    Timestamp  = $ts
                    Protocol   = 'IMAP4'
                    ClientIP   = $clientIP
                    Username   = $username
                    AuthMethod = 'LOGIN'
                    Success    = $result.Success
                    Error      = $result.Error
                    LogFile    = $file.Name
                })
            }

            # ── AUTHENTICATE command (SASL: PLAIN, LOGIN, NTLM, etc.) ────────
            elseif ($cmd -eq 'authenticate') {
                $mechanism = $params   # e.g. "PLAIN", "LOGIN", "NTLM"
                $username  = $userField   # populated by Exchange after SASL completes

                $result = Get-ContextResult $context
                $authEvents.Add([pscustomobject]@{
                    Timestamp  = $ts
                    Protocol   = 'IMAP4'
                    ClientIP   = $clientIP
                    Username   = $username
                    AuthMethod = "AUTH $mechanism"
                    Success    = $result.Success
                    Error      = $result.Error
                    LogFile    = $file.Name
                })
            }
        }
    }

    return $authEvents
}

#endregion

#region ── POP3 Log Parser ────────────────────────────────────────────────────

function Parse-Pop3Logs {
    <#
    Exchange POP3 log format (Exchange 2013/2016/2019):
      Fields: dateTime, sessionId, seqNumber, sIp, cIp, user,
              duration, rqsize, rpsize, command, parameters, context, puid

    Authentication flow:
      user command  - client sends username
                      parameters = the username sent
                      context = R=OK (server acknowledged)
      pass command  - client sends password (masked as *****)
                      context = R=OK;Msg="Proxy:...;ProxySuccess"   → success
                      context = R=<ErrorCode>;Msg="..."              → failure

    Session tracking is needed because username (from 'user' row) and
    auth result (from 'pass' row) are in separate log records.
    The 'user' field in the 'pass' row is truncated – always prefer
    the full username captured from the 'user' command row.
    #>
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

        # Per-session tracking: sessionId → { User, ClientIP, AuthTimestamp }
        $sessions = @{}

        try {
            $records = Parse-W3CLogFile -Path $file.FullName
        }
        catch {
            Write-Warning "Failed to read $($file.Name): $_"
            continue
        }

        foreach ($rec in $records) {
            # ── Timestamp ────────────────────────────────────────────────────
            $tsRaw = Get-RecordField $rec 'dateTime','date-time','DateTime'
            $ts    = [datetime]::MinValue
            if (-not $tsRaw -or -not [datetime]::TryParse($tsRaw, [ref]$ts)) { continue }
            if ($ts -lt $Start -or $ts -gt $End) { continue }

            # ── Fields ───────────────────────────────────────────────────────
            $sessionId = Get-RecordField $rec 'sessionId','session'
            $cmd       = (Get-RecordField $rec 'command').ToLower()
            $params    = Get-RecordField $rec 'parameters'
            $context   = Get-RecordField $rec 'context'
            $userField = Get-RecordField $rec 'user'
            $clientIP  = Extract-IPFromEndpoint (Get-RecordField $rec 'cIp','remote-endpoint')

            # Initialise session entry
            if ($sessionId -and -not $sessions.ContainsKey($sessionId)) {
                $sessions[$sessionId] = @{
                    User          = ''
                    ClientIP      = $clientIP
                    AuthTimestamp = $ts
                }
            }
            $sess = if ($sessionId) { $sessions[$sessionId] } else {
                @{ User = ''; ClientIP = $clientIP; AuthTimestamp = $ts }
            }
            if ($clientIP -ne '-') { $sess.ClientIP = $clientIP }

            switch ($cmd) {
                'user' {
                    # 'parameters' has the full username as typed by the client
                    # (could be domain\user, user@domain, or plain user)
                    $sess.User          = if ($params) { $params } else { $userField }
                    $sess.AuthTimestamp = $ts
                }

                'pass' {
                    # Auth result is in the context field of this very row.
                    # Use full username from session; fall back to (possibly truncated) user field.
                    $username = if ($sess.User) { $sess.User } else { $userField }
                    $result   = Get-ContextResult $context

                    if ($username) {
                        $authEvents.Add([pscustomobject]@{
                            Timestamp  = $sess.AuthTimestamp
                            Protocol   = 'POP3'
                            ClientIP   = $sess.ClientIP
                            Username   = $username
                            AuthMethod = 'USER/PASS'
                            Success    = $result.Success
                            Error      = $result.Error
                            LogFile    = $file.Name
                        })
                    }
                    $sess.User = ''
                }

                'auth' {
                    # SASL authentication (AUTH PLAIN, AUTH LOGIN, etc.)
                    # Username is in the user field after SASL completes.
                    $mechanism = $params
                    $username  = if ($sess.User) { $sess.User } else { $userField }
                    $result    = Get-ContextResult $context

                    if ($username -and $context) {
                        $authEvents.Add([pscustomobject]@{
                            Timestamp  = $ts
                            Protocol   = 'POP3'
                            ClientIP   = $sess.ClientIP
                            Username   = $username
                            AuthMethod = "AUTH $mechanism"
                            Success    = $result.Success
                            Error      = $result.Error
                            LogFile    = $file.Name
                        })
                    }
                    $sess.User = ''
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
Write-Host "[1/2] Processing IMAP4 logs..." -ForegroundColor Yellow
Write-Host "[2/2] Processing POP3 logs..." -ForegroundColor Yellow

$allEvents = @(
    Parse-ImapLogs -LogDirectory $ImapLogPath -Start $StartDate -End $EndDate
    Parse-Pop3Logs -LogDirectory $PopLogPath  -Start $StartDate -End $EndDate
) | Where-Object { $_ -ne $null }

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
