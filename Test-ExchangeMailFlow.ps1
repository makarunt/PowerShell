<#
.SYNOPSIS
    Pretražuje Exchange message tracking logove za provjeru outbound mail flowa.

.DESCRIPTION
    Skripta koristi Get-MessageTrackingLog za praćenje poruka kroz Exchange transport
    infrastrukturu. Prikazuje koji Exchange server je obradio poruku, kroz koji
    Send Connector je mail otišao i da li je odredišni sustav prihvatio poruku.

    Podržava pretragu po pošiljatelju, primatelju, predmetu i vremenskom rasponu.
    Skripta pretražuje sve Exchange transport servere u organizaciji.

.PARAMETER From
    Email adresa pošiljatelja (npr. korisnik@domena.hr).
    Podržava wildcard, npr. *@domena.hr

.PARAMETER To
    Email adresa primatelja (npr. vanjski@externaldomena.com).
    Podržava wildcard.

.PARAMETER Subject
    Predmet poruke. Podržava wildcard (npr. "*Testni mail*").

.PARAMETER MinutesBack
    Koliko minuta u prošlost pretražiti. Zadano: 15 minuta.
    Ako se skripta pokrene bez parametara pretrage (From/To/Subject), automatski
    se koristi 5 minuta.

.PARAMETER Servers
    Lista Exchange transport servera koje pretražujemo.
    Ako nije navedeno, automatski dohvaća sve transport servere u organizaciji.

.PARAMETER ShowAllEvents
    Prikaži sve događaje za pronađene poruke, ne samo ključne.

.EXAMPLE
    .\Test-ExchangeMailFlow.ps1 -From "korisnik@firma.hr" -To "vanjski@partner.com" -MinutesBack 30

.EXAMPLE
    .\Test-ExchangeMailFlow.ps1 -From "korisnik@firma.hr" -To "*@partner.com" -Subject "*Testni*" -MinutesBack 120

.EXAMPLE
    .\Test-ExchangeMailFlow.ps1 -From "*@firma.hr" -To "vanjski@partner.com" -MinutesBack 60 -ShowAllEvents

.NOTES
    Pokretati u Exchange Management Shell ili PowerShell sesiji s učitanim Exchange snap-inom.
    Potrebne dozvole: View-Only Organization Management ili Message Tracking role.
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

#region --- Inicijalizacija i provjere ---

# Fix za prikaz dijakritičkih znakova u konzoli
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Ključni eventi za outbound mail flow
$KeyOutboundEvents = @('RECEIVE', 'SEND', 'SENDEXTERNAL', 'FAIL', 'DEFER', 'REDIRECT', 'RESOLVE', 'TRANSFER', 'HADISCARD')

# Connector koji označava interni relay između Exchange servera (ne zanima nas kao "outbound")
$IntraOrgConnector = 'Intra-Organization SMTP Send Connector'

# Boje za ispis statusa
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

# Ako nema kriterija pretrage, prikaži sve iz zadnjih 5 minuta
$NoFilterMode = -not $From -and -not $To -and -not $Subject
if ($NoFilterMode) {
    # Ako korisnik nije eksplicitno postavio MinutesBack, koristi 5 minuta
    if (-not $PSBoundParameters.ContainsKey('MinutesBack')) {
        $MinutesBack = 5
    }
}

# Provjera dostupnosti Exchange cmdleta
if (-not (Get-Command Get-MessageTrackingLog -ErrorAction SilentlyContinue)) {
    Write-Error @"
Cmdlet Get-MessageTrackingLog nije dostupan.
Pokrenite skriptu iz Exchange Management Shell ili dodajte Exchange snap-in:
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
"@
    exit 1
}

#endregion

#region --- Dohvat Exchange servera ---

Write-Header "Exchange Mail Flow Tracker"

$StartTime = (Get-Date).AddMinutes(-$MinutesBack)
$EndTime   = Get-Date

Write-Host "`nKriteriji pretrage:" -ForegroundColor White
if ($NoFilterMode) {
    Write-Host "  Bez filtera - prikazujem sve poruke iz zadnjih $MinutesBack minuta" -ForegroundColor Yellow
} else {
    Write-Host "  Pošiljatelj  : $(if ($From)    { $From    } else { '(nije naveden)' })" -ForegroundColor Gray
    Write-Host "  Primatelj    : $(if ($To)      { $To      } else { '(nije naveden)' })" -ForegroundColor Gray
    Write-Host "  Predmet      : $(if ($Subject) { $Subject } else { '(nije naveden)' })" -ForegroundColor Gray
}
Write-Host "  Vremenski r. : $($StartTime.ToString('dd.MM.yyyy HH:mm:ss')) - $($EndTime.ToString('dd.MM.yyyy HH:mm:ss'))" -ForegroundColor Gray
Write-Host ""

if (-not $Servers) {
    Write-Host "Dohvaćam listu Exchange transport servera..." -ForegroundColor DarkGray
    try {
        $Servers = Get-TransportService | Select-Object -ExpandProperty Name
    }
    catch {
        Write-Warning "Ne mogu dohvatiti transport servere automatski. Koristim lokalni server."
        $Servers = @($env:COMPUTERNAME)
    }
}

Write-Host "Exchange serveri koji se pretražuju ($($Servers.Count)):" -ForegroundColor White
$Servers | ForEach-Object { Write-Host "  - $_" -ForegroundColor Gray }

#endregion

#region --- Pretraga message tracking logova ---

Write-SubHeader "Pretraga message tracking logova"

$TrackingParams = @{
    Start = $StartTime
    End   = $EndTime
}

if ($From)    { $TrackingParams['Sender']         = $From    }
if ($To)      { $TrackingParams['Recipients']     = $To      }
if ($Subject) { $TrackingParams['MessageSubject'] = $Subject }

$AllEvents = [System.Collections.Generic.List[object]]::new()
$ServerErrors = [System.Collections.Generic.List[string]]::new()

foreach ($Server in $Servers) {
    Write-Host "  Pretražujem: $Server" -ForegroundColor DarkGray -NoNewline
    try {
        $Events = Get-MessageTrackingLog @TrackingParams -Server $Server -ResultSize Unlimited -ErrorAction Stop
        $Count = ($Events | Measure-Object).Count
        Write-Host " -> $Count događaj(a)" -ForegroundColor DarkGray
        if ($Count -gt 0) {
            $Events | ForEach-Object {
                $_ | Add-Member -NotePropertyName 'TrackingServer' -NotePropertyValue $Server -Force
                $AllEvents.Add($_)
            }
        }
    }
    catch {
        Write-Host " -> GREŠKA: $($_.Exception.Message)" -ForegroundColor Red
        $ServerErrors.Add($Server)
    }
}

if ($AllEvents.Count -eq 0) {
    Write-Host "`nNisu pronađene poruke prema zadanim kriterijima." -ForegroundColor Yellow
    Write-Host "Savjeti:" -ForegroundColor DarkYellow
    Write-Host "  - Proširite vremenski raspon (-MinutesBack)" -ForegroundColor Gray
    Write-Host "  - Provjerite pravopis email adresa" -ForegroundColor Gray
    Write-Host "  - Koristite wildcard: -From '*@domena.hr'" -ForegroundColor Gray
    exit 0
}

#endregion

#region --- Grupiranje po poruci i analiza ---

Write-Header "Rezultati pretrage - ukupno $($AllEvents.Count) događaj(a)"

# Grupiramo po MessageId kako bi pratili svaku poruku zasebno
$MessageGroups = $AllEvents | Group-Object -Property MessageId | Where-Object {
    # Preskoči grupe koje imaju ISKLJUČIVO HADISCARD evente - to su shadow/DR log entriji
    $_.Group | Where-Object { $_.EventId -ne 'HADISCARD' }
} | Sort-Object {
    ($_.Group | Sort-Object Timestamp | Select-Object -First 1).Timestamp
} -Descending

Write-Host "Pronađeno $($MessageGroups.Count) jedinstvena(ih) poruka.`n" -ForegroundColor White

$MessageCounter = 0

foreach ($MessageGroup in $MessageGroups) {
    $MessageCounter++
    $GroupEvents = $MessageGroup.Group | Sort-Object Timestamp

    # Dohvati osnovne podatke o poruci
    $FirstEvent = $GroupEvents | Select-Object -First 1
    $LastEvent  = $GroupEvents | Select-Object -Last 1

    $MsgFrom    = $FirstEvent.Sender
    $MsgTo      = ($FirstEvent.Recipients -join ', ')
    $MsgSubject = $FirstEvent.MessageSubject
    $MsgId      = $FirstEvent.MessageId
    $MsgSize    = if ($FirstEvent.TotalBytes) { "$([math]::Round($FirstEvent.TotalBytes / 1KB, 1)) KB" } else { 'N/A' }

    # Odredi ukupni status poruke
    $HasFail         = $GroupEvents | Where-Object { $_.EventId -eq 'FAIL' }
    $HasDefer        = $GroupEvents | Where-Object { $_.EventId -eq 'DEFER' }
    $HasDeliver      = $GroupEvents | Where-Object { $_.EventId -eq 'DELIVER' }
    # Vanjski send: SENDEXTERNAL ili SEND koji NIJE intra-org connector
    $HasSendExternal = $GroupEvents | Where-Object {
        $_.EventId -eq 'SENDEXTERNAL' -or
        ($_.EventId -eq 'SEND' -and $_.ConnectorId -and $_.ConnectorId -notlike "*$IntraOrgConnector*")
    }
    # Sve SEND (uključujući intra-org), samo za fallback status
    $HasSend         = $GroupEvents | Where-Object { $_.EventId -eq 'SEND' }

    $OverallStatus = if ($HasFail)         { "FAIL (isporuka neuspjesna)" }
                     elseif ($HasDefer)         { "DEFER (privremeno odgodeno)" }
                     elseif ($HasSendExternal)  { "SENT EXTERNAL (poslano prema van)" }
                     elseif ($HasSend)          { "RELAYED (proslijedeno interno)" }
                     elseif ($HasDeliver)       { "DELIVERED (dostavljeno lokalno)" }
                     else                       { "IN PROGRESS (u obradi)" }

    $StatusColor = if ($HasFail)         { 'Red' }
                   elseif ($HasDefer)         { 'Yellow' }
                   elseif ($HasSendExternal)  { 'Green' }
                   elseif ($HasSend)          { 'DarkYellow' }
                   elseif ($HasDeliver)       { 'Green' }
                   else                       { 'Cyan' }

    # Naslov poruke
    $line = '-' * 70
    Write-Host "`n$line" -ForegroundColor DarkGray
    Write-Host "  PORUKA $MessageCounter/$($MessageGroups.Count)" -ForegroundColor White -NoNewline
    Write-Host "  [$OverallStatus]" -ForegroundColor $StatusColor
    Write-Host "$line" -ForegroundColor DarkGray

    Write-Host "  Od       : $MsgFrom" -ForegroundColor White
    Write-Host "  Za       : $MsgTo" -ForegroundColor White
    Write-Host "  Predmet  : $MsgSubject" -ForegroundColor White
    Write-Host "  Veličina : $MsgSize" -ForegroundColor Gray
    Write-Host "  Msg-ID   : $MsgId" -ForegroundColor DarkGray

    # --- Prikaz događaja ---
    Write-Host "`n  Tijek obrade:" -ForegroundColor White

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
        $EventColor    = Get-EventColor -EventId $Event.EventId
        $EventTime     = $Event.Timestamp.ToString('dd.MM.yyyy HH:mm:ss')
        $EventServer   = $Event.TrackingServer
        $EventId       = $Event.EventId.PadRight(10)
        $ConnectorId   = if ($Event.ConnectorId)   { $Event.ConnectorId }   else { '' }
        $SourceCtx     = if ($Event.SourceContext) { $Event.SourceContext } else { '' }
        $NextHop       = if ($Event.NextHopDomain) { $Event.NextHopDomain } else { '' }
        $NextHopConn   = if ($Event.NextHopConnector) { $Event.NextHopConnector } else { '' }

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

        # Detalji konteksta (odgovor odredišnog servera)
        if ($SourceCtx -and $Event.EventId -in @('SEND', 'SENDEXTERNAL', 'FAIL', 'DEFER')) {
            # Izvuci SMTP response kod ako postoji (bez ^ - nije nužno na početku stringa)
            if ($SourceCtx -match '(\d{3}\s.+?)(?:;|$)') {
                $SmtpResponse = $Matches[1].Trim()
                $RespColor = if ($SmtpResponse -match '^2\d\d') { 'Green' }
                             elseif ($SmtpResponse -match '^4\d\d') { 'Yellow' }
                             elseif ($SmtpResponse -match '^5\d\d') { 'Red' }
                             else { 'Gray' }
                Write-Host ("           SMTP odgovor: {0}" -f $SmtpResponse) -ForegroundColor $RespColor
            }
            else {
                Write-Host ("           Kontekst: {0}" -f ($SourceCtx -replace ';', '; ')) -ForegroundColor DarkGray
            }
        }

        # Upozorenje za FAIL
        if ($Event.EventId -eq 'FAIL') {
            Write-Host "  !! ISPORUKA NIJE USPJELA !!" -ForegroundColor Red
        }
    }

    # --- Sažetak vanjskog outbound hopa ---
    if ($HasSendExternal) {
        Write-Host "`n  Zadnji vanjski hop (izlaz iz organizacije):" -ForegroundColor White
        foreach ($SendEvent in $HasSendExternal) {
            $ConnInfo  = if ($SendEvent.ConnectorId) { $SendEvent.ConnectorId } else { 'N/A' }

            # Izvuci odredišni hostname iz SourceContext (Hostname= polje unutar SMTP odgovora)
            $DestHost  = if ($SendEvent.SourceContext -match 'Hostname=([^\],\s\[]+)') {
                             $Matches[1]
                         } elseif ($SendEvent.NextHopDomain) {
                             $SendEvent.NextHopDomain
                         } else { $null }

            # Ako nema hostname, izvuci domenu iz adrese primatelja kao ciljna domena
            $RecipDomain = $null
            if (-not $DestHost -and $SendEvent.Recipients) {
                $firstRecip = @($SendEvent.Recipients)[0]
                if ($firstRecip -match '@(.+)$') { $RecipDomain = $Matches[1] }
            }

            # Izvuci IP adresu ako postoji
            $RemoteIP  = if ($SendEvent.SourceContext -match 'RemoteEndpoint=\[?([0-9a-fA-F.:]+)\]?') {
                             $Matches[1]
                         } elseif ($SendEvent.SourceContext -match '\b(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})\b') {
                             $Matches[1]
                         } else { $null }

            # Odredi da li je odredišni server prihvatio poruku
            # Bez ^ - SMTP kod ne mora biti na samom početku SourceContext stringa
            $SmtpCode  = if ($SendEvent.SourceContext -match '(\d{3})\s') { $Matches[1] } else { $null }
            if ($SmtpCode -and $SmtpCode -notmatch '^[245]\d\d') { $SmtpCode = $null }

            $AcceptStatus = if ($SmtpCode -match '^2') {
                                "PRIHVACENO ($SmtpCode)"
                            } elseif ($SmtpCode -match '^4') {
                                "PRIVREMENO ODBIJENO ($SmtpCode)"
                            } elseif ($SmtpCode -match '^5') {
                                "TRAJNO ODBIJENO ($SmtpCode)"
                            } elseif (-not $SendEvent.SourceContext) {
                                "(relay preuzeo isporuku - provjeri relay logs)"
                            } else { 'nepoznato' }
            $AcceptColor = if ($SmtpCode -match '^2') { 'Green' }
                           elseif ($SmtpCode -match '^4') { 'Yellow' }
                           elseif ($SmtpCode -match '^5') { 'Red' }
                           else { 'Gray' }

            Write-Host ("    Exchange server     : {0}" -f $SendEvent.TrackingServer) -ForegroundColor Cyan
            Write-Host ("    Send Connector      : {0}" -f $ConnInfo) -ForegroundColor Magenta
            if ($DestHost) {
                Write-Host ("    Odredisni server    : {0}" -f $DestHost) -ForegroundColor Yellow
            } elseif ($RecipDomain) {
                Write-Host ("    Ciljna domena       : {0}" -f $RecipDomain) -ForegroundColor Yellow
            } else {
                Write-Host ("    Odredisni server    : N/A") -ForegroundColor DarkGray
            }
            if ($RemoteIP) {
                Write-Host ("    Remote IP           : {0}" -f $RemoteIP) -ForegroundColor Gray
            }
            Write-Host ("    Prihvacanje poruke  : ") -ForegroundColor White -NoNewline
            Write-Host $AcceptStatus -ForegroundColor $AcceptColor
            Write-Host ("    Timestamp           : {0}" -f $SendEvent.Timestamp.ToString('dd.MM.yyyy HH:mm:ss')) -ForegroundColor DarkGray
        }
    }
}

#endregion

#region --- Završni sažetak ---

Write-Header "Sažetak"

$TotalMessages        = $MessageGroups.Count
$SentExternalMessages = ($MessageGroups | Where-Object {
    $_.Group | Where-Object {
        $_.EventId -eq 'SENDEXTERNAL' -or
        ($_.EventId -eq 'SEND' -and $_.ConnectorId -and $_.ConnectorId -notlike "*$IntraOrgConnector*")
    }
}).Count
$RelayedMessages  = ($MessageGroups | Where-Object {
    ($_.Group.EventId -contains 'SEND') -and -not (
        $_.Group | Where-Object {
            $_.EventId -eq 'SENDEXTERNAL' -or
            ($_.EventId -eq 'SEND' -and $_.ConnectorId -and $_.ConnectorId -notlike "*$IntraOrgConnector*")
        }
    )
}).Count
$FailedMessages   = ($MessageGroups | Where-Object { $_.Group.EventId -contains 'FAIL' }).Count
$DeferMessages    = ($MessageGroups | Where-Object { $_.Group.EventId -contains 'DEFER' }).Count

Write-Host ("  Ukupno pronadenih poruka      : {0}" -f $TotalMessages) -ForegroundColor White
Write-Host ("  Poslano prema van (eksterno)  : {0}" -f $SentExternalMessages) -ForegroundColor Green
Write-Host ("  Intra-org relay (samo interno): {0}" -f $RelayedMessages) -ForegroundColor DarkYellow
Write-Host ("  Neuspjesnih (FAIL)            : {0}" -f $FailedMessages) -ForegroundColor $(if ($FailedMessages -gt 0) { 'Red' } else { 'Gray' })
Write-Host ("  Odgodenih (DEFER)             : {0}" -f $DeferMessages) -ForegroundColor $(if ($DeferMessages -gt 0) { 'Yellow' } else { 'Gray' })

if ($ServerErrors.Count -gt 0) {
    Write-Host "`n  Serveri s greškama (nisu pretraženi):" -ForegroundColor Red
    $ServerErrors | ForEach-Object { Write-Host "    - $_" -ForegroundColor Red }
}

Write-Host "`n  Pretraga završena: $(Get-Date -Format 'dd.MM.yyyy HH:mm:ss')" -ForegroundColor DarkGray
Write-Host ""

#endregion
