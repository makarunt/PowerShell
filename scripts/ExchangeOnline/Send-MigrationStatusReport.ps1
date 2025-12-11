<#
.SYNOPSIS
    Šalje email izvještaj o statusu mailbox migracija.

.DESCRIPTION
    Automatski generiše i šalje email izvještaj sa trenutnim statusom svih migration batch-eva.
    Idealno za scheduling sa Task Schedulerom za redovno izvještavanje.

.PARAMETER BatchName
    Naziv specifičnog migration batcha za izvještaj.

.PARAMETER To
    Email adresa(e) primaoca izvještaja (odvojene zarezom).

.PARAMETER From
    Email adresa pošiljaoca (default: noreply@<tenant>.onmicrosoft.com).

.PARAMETER SmtpServer
    SMTP server za slanje email-a (default: smtp.office365.com).

.PARAMETER IncludeOnlyReady
    Uključuje u izvještaj samo mailboxeve spremne za finalizaciju (≥95%).

.EXAMPLE
    .\Send-MigrationStatusReport.ps1 -To "admin@contoso.com"

    Šalje izvještaj o svim migracijama na navedenu adresu.

.EXAMPLE
    .\Send-MigrationStatusReport.ps1 -BatchName "Batch-Finance" -To "it-team@contoso.com" -IncludeOnlyReady

    Šalje izvještaj samo o mailboxevima spremnim za finalizaciju.

.NOTES
    Autor: PowerShell Migration Script
    Verzija: 1.0
    Datum: 2025-12-11

    Za slanje email-a koristi Send-MailMessage cmdlet koji zahtijeva SMTP pristup.
    Alternativa: Koristite Graph API za slanje putem Microsoft Graph.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BatchName,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string[]]$To,

    [Parameter(Mandatory = $false)]
    [string]$From,

    [Parameter(Mandatory = $false)]
    [string]$SmtpServer = "smtp.office365.com",

    [Parameter(Mandatory = $false)]
    [switch]$IncludeOnlyReady
)

#Requires -Modules ExchangeOnlineManagement

function New-MigrationHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$MigrationData,

        [Parameter(Mandatory = $false)]
        [switch]$OnlyReady
    )

    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

    # Filter data ako je traženo samo ready
    if ($OnlyReady) {
        $MigrationData = $MigrationData | Where-Object { $_.PercentageComplete -ge 95 }
    }

    # Grupiraj po batch-u
    $batchGroups = $MigrationData | Group-Object -Property BatchId

    # HTML stilovi
    $htmlStyle = @"
<style>
    body {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background-color: #f5f5f5;
        margin: 0;
        padding: 20px;
    }
    .container {
        max-width: 1200px;
        margin: 0 auto;
        background-color: white;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }
    h1 {
        color: #0078d4;
        border-bottom: 3px solid #0078d4;
        padding-bottom: 10px;
    }
    h2 {
        color: #106ebe;
        margin-top: 30px;
        border-left: 5px solid #0078d4;
        padding-left: 15px;
    }
    .summary {
        background-color: #f0f8ff;
        padding: 20px;
        border-radius: 8px;
        margin: 20px 0;
        border-left: 5px solid #0078d4;
    }
    .summary-item {
        font-size: 18px;
        margin: 10px 0;
    }
    .ready {
        color: #107c10;
        font-weight: bold;
    }
    .in-progress {
        color: #0078d4;
    }
    .failed {
        color: #d13438;
        font-weight: bold;
    }
    table {
        width: 100%;
        border-collapse: collapse;
        margin: 20px 0;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    th {
        background-color: #0078d4;
        color: white;
        padding: 12px;
        text-align: left;
        font-weight: 600;
    }
    td {
        padding: 10px;
        border-bottom: 1px solid #e1e1e1;
    }
    tr:hover {
        background-color: #f5f5f5;
    }
    .status-badge {
        padding: 5px 10px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: bold;
        display: inline-block;
    }
    .status-synced {
        background-color: #107c10;
        color: white;
    }
    .status-syncing {
        background-color: #0078d4;
        color: white;
    }
    .status-failed {
        background-color: #d13438;
        color: white;
    }
    .status-queued {
        background-color: #ffb900;
        color: black;
    }
    .progress-bar {
        width: 100%;
        height: 20px;
        background-color: #e1e1e1;
        border-radius: 10px;
        overflow: hidden;
    }
    .progress-fill {
        height: 100%;
        background: linear-gradient(90deg, #0078d4, #106ebe);
        text-align: center;
        line-height: 20px;
        color: white;
        font-size: 12px;
        font-weight: bold;
    }
    .footer {
        margin-top: 40px;
        padding-top: 20px;
        border-top: 1px solid #e1e1e1;
        color: #666;
        font-size: 12px;
        text-align: center;
    }
</style>
"@

    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    $htmlStyle
</head>
<body>
    <div class="container">
        <h1>📊 Exchange Online Mailbox Migration - Status izvještaj</h1>
        <p><strong>Datum/Vrijeme:</strong> $timestamp</p>

        <div class="summary">
            <div class="summary-item">📦 <strong>Ukupno batch-eva:</strong> $($batchGroups.Count)</div>
            <div class="summary-item">📧 <strong>Ukupno mailboxeva:</strong> $($MigrationData.Count)</div>
            <div class="summary-item ready">✅ <strong>Spremno za finalizaciju (≥95%):</strong> $(($MigrationData | Where-Object { $_.PercentageComplete -ge 95 }).Count)</div>
            <div class="summary-item in-progress">⏳ <strong>U toku (&lt;95%):</strong> $(($MigrationData | Where-Object { $_.PercentageComplete -lt 95 -and $_.Status -eq 'Syncing' }).Count)</div>
            <div class="summary-item failed">❌ <strong>Neuspjelo:</strong> $(($MigrationData | Where-Object { $_.Status -eq 'Failed' }).Count)</div>
        </div>
"@

    foreach ($batchGroup in $batchGroups) {
        $batchName = $batchGroup.Name
        $mailboxes = $batchGroup.Group | Sort-Object -Property PercentageComplete -Descending

        $readyCount = ($mailboxes | Where-Object { $_.PercentageComplete -ge 95 }).Count
        $readyIndicator = if ($readyCount -eq $mailboxes.Count) {
            "✅ SVE SPREMNO ZA FINALIZACIJU"
        }
        elseif ($readyCount -gt 0) {
            "⚠️ DJELOMIČNO SPREMNO ($readyCount/$($mailboxes.Count))"
        }
        else {
            "⏳ U TOKU"
        }

        $htmlBody += @"
        <h2>Batch: $batchName <span style="font-size: 14px; font-weight: normal;">$readyIndicator</span></h2>
        <table>
            <thead>
                <tr>
                    <th>Mailbox</th>
                    <th>Status</th>
                    <th>Progress</th>
                    <th>Trajanje</th>
                    <th>Bytes Transferred</th>
                </tr>
            </thead>
            <tbody>
"@

        foreach ($mailbox in $mailboxes) {
            $statusClass = switch ($mailbox.Status) {
                'Synced' { 'status-synced' }
                'Syncing' { 'status-syncing' }
                'Failed' { 'status-failed' }
                'Queued' { 'status-queued' }
                default { 'status-syncing' }
            }

            $duration = if ($mailbox.SyncDuration) {
                "{0:N0}d {1:N0}h {2:N0}m" -f $mailbox.SyncDuration.Days, $mailbox.SyncDuration.Hours, $mailbox.SyncDuration.Minutes
            }
            else {
                "N/A"
            }

            $bytes = if ($mailbox.BytesTransferred) {
                "{0:N2} GB" -f ($mailbox.BytesTransferred / 1GB)
            }
            else {
                "N/A"
            }

            $progressBarColor = if ($mailbox.PercentageComplete -ge 95) { '#107c10' } else { '#0078d4' }

            $htmlBody += @"
                <tr>
                    <td><strong>$($mailbox.Identity)</strong></td>
                    <td><span class="status-badge $statusClass">$($mailbox.Status)</span></td>
                    <td>
                        <div class="progress-bar">
                            <div class="progress-fill" style="width: $($mailbox.PercentageComplete)%; background: $progressBarColor;">
                                $($mailbox.PercentageComplete)%
                            </div>
                        </div>
                    </td>
                    <td>$duration</td>
                    <td>$bytes</td>
                </tr>
"@
        }

        $htmlBody += @"
            </tbody>
        </table>
"@
    }

    $htmlBody += @"
        <div class="footer">
            <p>Automatski generiran izvještaj - PowerShell Migration Scripts</p>
            <p>Za finalizaciju koristite: <code>Complete-MigrationBatch -Identity "BatchName"</code></p>
        </div>
    </div>
</body>
</html>
"@

    return $htmlBody
}

# Glavna skripta
try {
    # Provjera konekcije
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
    }
    catch {
        Write-Error "Nije uspostavljena konekcija sa Exchange Online. Pokrenite 'Connect-ExchangeOnline'."
        return
    }

    Write-Host "Prikupljam podatke o migracijama..." -ForegroundColor Cyan

    # Dohvati batch-eve
    if ($BatchName) {
        $batches = @(Get-MigrationBatch -Identity $BatchName -ErrorAction Stop)
    }
    else {
        $batches = @(Get-MigrationBatch -ErrorAction Stop | Where-Object { $_.Status -ne 'Completed' })
    }

    if ($batches.Count -eq 0) {
        Write-Warning "Nisu pronađeni migration batch-evi."
        return
    }

    # Prikupi sve podatke
    $allMigrationData = @()

    foreach ($batch in $batches) {
        $users = @(Get-MigrationUser -BatchId $batch.Identity -ErrorAction SilentlyContinue)

        foreach ($user in $users) {
            $stats = Get-MigrationUserStatistics -Identity $user.Identity -ErrorAction SilentlyContinue

            if ($stats) {
                $syncDuration = if ($stats.InitialSyncDateTime -and $stats.LastSyncedDateTime) {
                    $stats.LastSyncedDateTime - $stats.InitialSyncDateTime
                }
                else {
                    $null
                }

                $allMigrationData += [PSCustomObject]@{
                    Identity           = $stats.Identity
                    BatchId            = $stats.BatchId
                    Status             = $stats.Status
                    PercentageComplete = $stats.PercentageComplete
                    BytesTransferred   = $stats.BytesTransferred
                    SyncDuration       = $syncDuration
                }
            }
        }
    }

    if ($allMigrationData.Count -eq 0) {
        Write-Warning "Nisu pronađeni migration users."
        return
    }

    Write-Host "Generišem HTML izvještaj..." -ForegroundColor Cyan

    # Generiši HTML izvještaj
    $htmlReport = New-MigrationHtmlReport -MigrationData $allMigrationData -OnlyReady:$IncludeOnlyReady

    # Odredi From adresu ako nije navedena
    if (-not $From) {
        $orgConfig = Get-OrganizationConfig
        $tenantDomain = $orgConfig.Name
        $From = "noreply@$tenantDomain"
    }

    # Subject line
    $readyCount = ($allMigrationData | Where-Object { $_.PercentageComplete -ge 95 }).Count
    $subject = if ($IncludeOnlyReady) {
        "Migration Status: $readyCount mailbox(eva) spremno za finalizaciju"
    }
    else {
        "Migration Status: $($allMigrationData.Count) mailbox(eva) - $readyCount spremno"
    }

    Write-Host "Šaljem email izvještaj na $($To -join ', ')..." -ForegroundColor Cyan

    # Napomena: Send-MailMessage je deprecated, ali još uvijek funkcionalan
    # Za produkciju preporučujem korištenje Microsoft Graph API-ja

    try {
        # Probaj sa Send-MailMessage (zahtijeva SMTP pristup)
        $mailParams = @{
            To         = $To
            From       = $From
            Subject    = $subject
            Body       = $htmlReport
            BodyAsHtml = $true
            SmtpServer = $SmtpServer
            Port       = 587
            UseSsl     = $true
        }

        # Napomena: Možda će biti potrebna autentifikacija
        # $credential = Get-Credential
        # $mailParams['Credential'] = $credential

        Send-MailMessage @mailParams -ErrorAction Stop

        Write-Host "✅ Email izvještaj uspješno poslan!" -ForegroundColor Green
    }
    catch {
        Write-Warning "Nije moguće poslati email putem SMTP: $_"
        Write-Host "`n💡 Alternativa: Sačuvajte HTML izvještaj lokalno:" -ForegroundColor Cyan

        $reportPath = ".\MigrationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        $htmlReport | Out-File -FilePath $reportPath -Encoding UTF8

        Write-Host "   Izvještaj sačuvan u: $reportPath" -ForegroundColor Green
        Write-Host "   Možete ga poslati ručno ili koristiti Graph API za slanje." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Greška: $_"
    Write-Error $_.Exception.Message
}
