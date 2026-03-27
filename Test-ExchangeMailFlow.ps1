<#
.SYNOPSIS
    Searches Exchange message tracking logs to verify outbound mail flow.

.DESCRIPTION
    Uses Get-MessageTrackingLog to trace messages through the Exchange transport
    infrastructure. Shows which Exchange server processed the message, which
    Send Connector was used, and whether the destination system accepted the message.

    Supports filtering by sender, recipient, subject and time range.
    Automatically queries all Exchange transport servers in the organization.

.PARAMETER From
    Sender email address (e.g. user@domain.com).
    Supports wildcards, e.g. *@domain.com

.PARAMETER To
    Recipient email address (e.g. external@partner.com).
    Supports wildcards.

.PARAMETER Subject
    Message subject. Supports wildcards (e.g. "*Test mail*").

.PARAMETER MinutesBack
    How many minutes back to search. Default: 15 minutes.
    If the script is run without any search parameters (From/To/Subject),
    defaults to 5 minutes.

.PARAMETER Servers
    List of Exchange transport servers to query.
    If omitted, automatically discovers all transport servers in the organization.

.PARAMETER ShowAllEvents
    Show all events for found messages, not just the key ones.

.EXAMPLE
    .\Test-ExchangeMailFlow.ps1 -From "user@company.com" -To "external@partner.com" -MinutesBack 30

.EXAMPLE
    .\Test-ExchangeMailFlow.ps1 -From "user@company.com" -To "*@partner.com" -Subject "*Test*" -MinutesBack 120

.EXAMPLE
    .\Test-ExchangeMailFlow.ps1 -From "*@company.com" -To "external@partner.com" -MinutesBack 60 -ShowAllEvents

.NOTES
    Run from Exchange Management Shell or a PowerShell session with the Exchange snap-in loaded.
    Required permissions: View-Only Organization Management or Message Tracking role.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$From,

    [Parameter(Mandatory = $false)]
    [string]$To,

    [Parameter(Mandatory = $false)]
    [string]$Subject,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10080)]
    [int]$MinutesBack = 15,

    [Parameter(Mandatory = $false)]
    [string[]]$Servers,

    [Parameter(Mandatory = $false)]
    [switch]$ShowAllEvents
)

#region --- Initialization ---

# Ensure correct character encoding in the console
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Key events for outbound mail flow
$KeyOutboundEvents = @('RECEIVE', 'SEND', 'SENDEXTERNAL', 'FAIL', 'DEFER', 'REDIRECT', 'RESOLVE', 'TRANSFER', 'HADISCARD')

# Connector name that indicates internal relay between Exchange servers (not considered outbound)
$IntraOrgConnector = 'Intra-Organization SMTP Send Connector'

# Colors for event display
$StatusColors = @{
    'SEND'         = 'Green'
    'SENDEXTERNAL' = 'Green'
    'RECEIVE'      = 'Cyan'
    'DELIVER'      = 'Green'
    'FAIL'         = 'Red'
    'DEFER'        = 'Yellow'
    'REDIRECT'     = 'Magenta'
    'RESOLVE'      = 'Gray'
    'TRANSFER'     = 'Yellow'
    'HADISCARD'    = 'DarkGray'
    'DEFAULT'      = 'White'
}

function Write-Header {
    param([string]$Title)
    $line = '=' * 70
    Write-Host "`n$line" -ForegroundColor DarkCyan
    Write-Host "  $Title" -ForegroundColor Cyan
    Write-Host "$line" -ForegroundColor DarkCyan
}

function Write-SubHeader {
    param([string]$Title)
    Write-Host "`n--- $Title ---" -ForegroundColor DarkYellow
}

function Get-EventColor {
    param([string]$EventId)
    if ($StatusColors.ContainsKey($EventId)) {
        return $StatusColors[$EventId]
    }
    return $StatusColors['DEFAULT']
}

# If no search filters provided, show all messages from the last 5 minutes
$NoFilterMode = -not $From -and -not $To -and -not $Subject
if ($NoFilterMode) {
    if (-not $PSBoundParameters.ContainsKey('MinutesBack')) {
        $MinutesBack = 5
    }
}

# Check that the Exchange cmdlet is available
if (-not (Get-Command Get-MessageTrackingLog -ErrorAction SilentlyContinue)) {
    Write-Error @"
Get-MessageTrackingLog cmdlet is not available.
Run the script from Exchange Management Shell or load the Exchange snap-in:
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
"@
    exit 1
}

#endregion

#region --- Server discovery ---

Write-Header "Exchange Mail Flow Tracker"

$StartTime = (Get-Date).AddMinutes(-$MinutesBack)
$EndTime   = Get-Date

Write-Host "`nSearch criteria:" -ForegroundColor White
if ($NoFilterMode) {
    Write-Host "  No filter - showing all messages from the last $MinutesBack minutes" -ForegroundColor Yellow
} else {
    Write-Host "  From    : $(if ($From)    { $From    } else { '(not specified)' })" -ForegroundColor Gray
    Write-Host "  To      : $(if ($To)      { $To      } else { '(not specified)' })" -ForegroundColor Gray
    Write-Host "  Subject : $(if ($Subject) { $Subject } else { '(not specified)' })" -ForegroundColor Gray
}
Write-Host "  Range   : $($StartTime.ToString('dd.MM.yyyy HH:mm:ss')) - $($EndTime.ToString('dd.MM.yyyy HH:mm:ss'))" -ForegroundColor Gray
Write-Host ""

if (-not $Servers) {
    Write-Host "Discovering Exchange transport servers..." -ForegroundColor DarkGray
    try {
        $Servers = Get-TransportService | Select-Object -ExpandProperty Name
    }
    catch {
        Write-Warning "Could not retrieve transport servers automatically. Using local server."
        $Servers = @($env:COMPUTERNAME)
    }
}

Write-Host "Exchange servers to query ($($Servers.Count)):" -ForegroundColor White
$Servers | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

#endregion

#region --- Message tracking log search ---

Write-SubHeader "Searching message tracking logs"

$TrackingParams = @{
    Start = $StartTime
    End   = $EndTime
}

if ($From)    { $TrackingParams['Sender']         = $From    }
if ($To)      { $TrackingParams['Recipients']     = $To      }
if ($Subject) { $TrackingParams['MessageSubject'] = $Subject }

$AllEvents    = [System.Collections.Generic.List[object]]::new()
$ServerErrors = [System.Collections.Generic.List[string]]::new()

foreach ($Server in $Servers) {
    Write-Host "  Querying: $Server" -ForegroundColor DarkGray -NoNewline
    try {
        $Events = Get-MessageTrackingLog @TrackingParams -Server $Server -ResultSize Unlimited -ErrorAction Stop
        $Count = ($Events | Measure-Object).Count
        Write-Host " -> $Count event(s)" -ForegroundColor DarkGray
        if ($Count -gt 0) {
            $Events | ForEach-Object {
                $_ | Add-Member -NotePropertyName 'TrackingServer' -NotePropertyValue $Server -Force
                $AllEvents.Add($_)
            }
        }
    }
    catch {
        Write-Host " -> ERROR: $($_.Exception.Message)" -ForegroundColor Red
        $ServerErrors.Add($Server)
    }
}

if ($AllEvents.Count -eq 0) {
    Write-Host "`nNo messages found matching the specified criteria." -ForegroundColor Yellow
    Write-Host "Tips:" -ForegroundColor DarkYellow
    Write-Host "  - Increase the time range (-MinutesBack)" -ForegroundColor Gray
    Write-Host "  - Check the spelling of email addresses" -ForegroundColor Gray
    Write-Host "  - Use wildcards: -From '*@domain.com'" -ForegroundColor Gray
    exit 0
}

#endregion

#region --- Grouping and analysis ---

Write-Header "Search results - $($AllEvents.Count) event(s) found"

# Group by MessageId to track each message individually
$MessageGroups = $AllEvents | Group-Object -Property MessageId | Where-Object {
    # Skip groups that contain ONLY HADISCARD events - these are shadow/DR log entries
    $_.Group | Where-Object { $_.EventId -ne 'HADISCARD' }
} | Sort-Object {
    ($_.Group | Sort-Object Timestamp | Select-Object -First 1).Timestamp
} -Descending

Write-Host "Found $($MessageGroups.Count) unique message(s).`n" -ForegroundColor White

$MessageCounter = 0

foreach ($MessageGroup in $MessageGroups) {
    $MessageCounter++
    $GroupEvents = $MessageGroup.Group | Sort-Object Timestamp

    # Get basic message info
    $FirstEvent = $GroupEvents | Select-Object -First 1

    $MsgFrom    = $FirstEvent.Sender
    $MsgTo      = ($FirstEvent.Recipients -join ', ')
    $MsgSubject = $FirstEvent.MessageSubject
    $MsgId      = $FirstEvent.MessageId
    $MsgSize    = if ($FirstEvent.TotalBytes) { "$([math]::Round($FirstEvent.TotalBytes / 1KB, 1)) KB" } else { 'N/A' }

    # Determine overall message status
    $HasFail         = $GroupEvents | Where-Object { $_.EventId -eq 'FAIL' }
    $HasDefer        = $GroupEvents | Where-Object { $_.EventId -eq 'DEFER' }
    $HasDeliver      = $GroupEvents | Where-Object { $_.EventId -eq 'DELIVER' }
    # External send: SENDEXTERNAL or SEND that is NOT via intra-org connector
    $HasSendExternal = $GroupEvents | Where-Object {
        $_.EventId -eq 'SENDEXTERNAL' -or
        ($_.EventId -eq 'SEND' -and $_.ConnectorId -and $_.ConnectorId -notlike "*$IntraOrgConnector*")
    }
    # All SEND events including intra-org, used as fallback status only
    $HasSend         = $GroupEvents | Where-Object { $_.EventId -eq 'SEND' }

    $OverallStatus = if ($HasFail)         { "FAIL (delivery failed)" }
                     elseif ($HasDefer)         { "DEFER (temporarily deferred)" }
                     elseif ($HasSendExternal)  { "SENT EXTERNAL (delivered to external server)" }
                     elseif ($HasSend)          { "RELAYED (internal relay only)" }
                     elseif ($HasDeliver)       { "DELIVERED (delivered to local mailbox)" }
                     else                       { "IN PROGRESS" }

    $StatusColor = if ($HasFail)         { 'Red' }
                   elseif ($HasDefer)         { 'Yellow' }
                   elseif ($HasSendExternal)  { 'Green' }
                   elseif ($HasSend)          { 'DarkYellow' }
                   elseif ($HasDeliver)       { 'Green' }
                   else                       { 'Cyan' }

    # Message header
    $line = '-' * 70
    Write-Host "`n$line" -ForegroundColor DarkGray
    Write-Host "  MESSAGE $MessageCounter/$($MessageGroups.Count)" -ForegroundColor White -NoNewline
    Write-Host "  [$OverallStatus]" -ForegroundColor $StatusColor
    Write-Host "$line" -ForegroundColor DarkGray

    Write-Host "  From    : $MsgFrom" -ForegroundColor White
    Write-Host "  To      : $MsgTo" -ForegroundColor White
    Write-Host "  Subject : $MsgSubject" -ForegroundColor White
    Write-Host "  Size    : $MsgSize" -ForegroundColor Gray
    Write-Host "  Msg-ID  : $MsgId" -ForegroundColor DarkGray

    # --- Event timeline ---
    Write-Host "`n  Processing timeline:" -ForegroundColor White

    $EventsToShow = if ($ShowAllEvents) {
        $GroupEvents
    }
    else {
        $GroupEvents | Where-Object { $_.EventId -in $KeyOutboundEvents }
    }

    if (-not $EventsToShow) {
        $EventsToShow = $GroupEvents
    }

    foreach ($Event in $EventsToShow) {
        $EventColor  = Get-EventColor -EventId $Event.EventId
        $EventTime   = $Event.Timestamp.ToString('dd.MM.yyyy HH:mm:ss')
        $EventServer = $Event.TrackingServer
        $EventId     = $Event.EventId.PadRight(14)
        $ConnectorId = if ($Event.ConnectorId)      { $Event.ConnectorId }      else { '' }
        $SourceCtx   = if ($Event.SourceContext)    { $Event.SourceContext }    else { '' }
        $NextHop     = if ($Event.NextHopDomain)    { $Event.NextHopDomain }    else { '' }

        Write-Host ("  {0}  " -f $EventTime) -ForegroundColor DarkGray -NoNewline
        Write-Host ("{0}" -f $EventId) -ForegroundColor $EventColor -NoNewline
        Write-Host ("  [{0}]" -f $EventServer) -ForegroundColor Cyan -NoNewline

        if ($ConnectorId) {
            Write-Host ("  Connector: {0}" -f $ConnectorId) -ForegroundColor Magenta -NoNewline
        }
        if ($NextHop) {
            Write-Host ("  NextHop: {0}" -f $NextHop) -ForegroundColor Yellow -NoNewline
        }
        Write-Host ""

        # Show SMTP response context for relevant events
        if ($SourceCtx -and $Event.EventId -in @('SEND', 'SENDEXTERNAL', 'FAIL', 'DEFER')) {
            if ($SourceCtx -match '(\d{3}\s.+?)(?:;|$)') {
                $SmtpResponse = $Matches[1].Trim()
                $RespColor = if ($SmtpResponse -match '^2\d\d') { 'Green' }
                             elseif ($SmtpResponse -match '^4\d\d') { 'Yellow' }
                             elseif ($SmtpResponse -match '^5\d\d') { 'Red' }
                             else { 'Gray' }
                Write-Host ("             SMTP response: {0}" -f $SmtpResponse) -ForegroundColor $RespColor
            }
            else {
                Write-Host ("             Context: {0}" -f ($SourceCtx -replace ';', '; ')) -ForegroundColor DarkGray
            }
        }

        if ($Event.EventId -eq 'FAIL') {
            Write-Host "  !! DELIVERY FAILED !!" -ForegroundColor Red
        }
    }

    # --- External outbound hop summary ---
    if ($HasSendExternal) {
        Write-Host "`n  Last external hop (exit from organization):" -ForegroundColor White
        foreach ($SendEvent in $HasSendExternal) {
            $ConnInfo = if ($SendEvent.ConnectorId) { $SendEvent.ConnectorId } else { 'N/A' }

            # Extract destination hostname from SourceContext (Hostname= field in SMTP response)
            $DestHost = if ($SendEvent.SourceContext -match 'Hostname=([^\],\s\[]+)') {
                            $Matches[1]
                        } elseif ($SendEvent.NextHopDomain) {
                            $SendEvent.NextHopDomain
                        } else { $null }

            # If no hostname available, extract recipient domain as fallback
            $RecipDomain = $null
            if (-not $DestHost -and $SendEvent.Recipients) {
                $firstRecip = @($SendEvent.Recipients)[0]
                if ($firstRecip -match '@(.+)$') { $RecipDomain = $Matches[1] }
            }

            # Extract remote IP if available
            $RemoteIP = if ($SendEvent.SourceContext -match 'RemoteEndpoint=\[?([0-9a-fA-F.:]+)\]?') {
                            $Matches[1]
                        } elseif ($SendEvent.SourceContext -match '\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b') {
                            $Matches[1]
                        } else { $null }

            # Determine whether the destination server accepted the message
            $SmtpCode = if ($SendEvent.SourceContext -match '(\d{3})\s') { $Matches[1] } else { $null }
            if ($SmtpCode -and $SmtpCode -notmatch '^[245]\d\d') { $SmtpCode = $null }

            $AcceptStatus = if ($SmtpCode -match '^2') {
                                "ACCEPTED ($SmtpCode)"
                            } elseif ($SmtpCode -match '^4') {
                                "TEMPORARILY REJECTED ($SmtpCode)"
                            } elseif ($SmtpCode -match '^5') {
                                "PERMANENTLY REJECTED ($SmtpCode)"
                            } elseif (-not $SendEvent.SourceContext) {
                                "(relay accepted handoff - check relay logs for final delivery)"
                            } else { 'unknown' }
            $AcceptColor = if ($SmtpCode -match '^2') { 'Green' }
                           elseif ($SmtpCode -match '^4') { 'Yellow' }
                           elseif ($SmtpCode -match '^5') { 'Red' }
                           else { 'Gray' }

            Write-Host ("    Exchange server  : {0}" -f $SendEvent.TrackingServer) -ForegroundColor Cyan
            Write-Host ("    Send Connector   : {0}" -f $ConnInfo) -ForegroundColor Magenta
            if ($DestHost) {
                Write-Host ("    Destination host : {0}" -f $DestHost) -ForegroundColor Yellow
            } elseif ($RecipDomain) {
                Write-Host ("    Recipient domain : {0}" -f $RecipDomain) -ForegroundColor Yellow
            } else {
                Write-Host ("    Destination host : N/A") -ForegroundColor DarkGray
            }
            if ($RemoteIP) {
                Write-Host ("    Remote IP        : {0}" -f $RemoteIP) -ForegroundColor Gray
            }
            Write-Host ("    Acceptance       : ") -ForegroundColor White -NoNewline
            Write-Host $AcceptStatus -ForegroundColor $AcceptColor
            Write-Host ("    Timestamp        : {0}" -f $SendEvent.Timestamp.ToString('dd.MM.yyyy HH:mm:ss')) -ForegroundColor DarkGray
        }
    }
}

#endregion

#region --- Summary ---

Write-Header "Summary"

$TotalMessages        = $MessageGroups.Count
$SentExternalMessages = ($MessageGroups | Where-Object {
    $_.Group | Where-Object {
        $_.EventId -eq 'SENDEXTERNAL' -or
        ($_.EventId -eq 'SEND' -and $_.ConnectorId -and $_.ConnectorId -notlike "*$IntraOrgConnector*")
    }
}).Count
$RelayedMessages = ($MessageGroups | Where-Object {
    ($_.Group.EventId -contains 'SEND') -and -not (
        $_.Group | Where-Object {
            $_.EventId -eq 'SENDEXTERNAL' -or
            ($_.EventId -eq 'SEND' -and $_.ConnectorId -and $_.ConnectorId -notlike "*$IntraOrgConnector*")
        }
    )
}).Count
$FailedMessages  = ($MessageGroups | Where-Object { $_.Group.EventId -contains 'FAIL' }).Count
$DeferMessages   = ($MessageGroups | Where-Object { $_.Group.EventId -contains 'DEFER' }).Count

Write-Host ("  Total messages found          : {0}" -f $TotalMessages) -ForegroundColor White
Write-Host ("  Sent external                 : {0}" -f $SentExternalMessages) -ForegroundColor Green
Write-Host ("  Internal relay only           : {0}" -f $RelayedMessages) -ForegroundColor DarkYellow
Write-Host ("  Failed (FAIL)                 : {0}" -f $FailedMessages) -ForegroundColor $(if ($FailedMessages -gt 0) { 'Red' } else { 'Gray' })
Write-Host ("  Deferred (DEFER)              : {0}" -f $DeferMessages) -ForegroundColor $(if ($DeferMessages -gt 0) { 'Yellow' } else { 'Gray' })

if ($ServerErrors.Count -gt 0) {
    Write-Host "`n  Servers with errors (not queried):" -ForegroundColor Red
    $ServerErrors | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
}

Write-Host "`n  Search completed: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')" -ForegroundColor DarkGray
Write-Host ""

#endregion
