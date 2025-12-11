<#
.SYNOPSIS
    Salje email izvjestaj o statusu mailbox migracija.

.DESCRIPTION
    Automatski generise i salje email izvjestaj sa trenutnim statusom svih migration batch-eva.
    Idealno za scheduling sa Task Schedulerom za redovno izvjestavanj e.

.PARAMETER BatchName
    Naziv specificnog migration batcha za izvjestaj.

.PARAMETER To
    Email adresa(e) primaoca izvjestaja (odvojene zarezom).

.PARAMETER From
    Email adresa posiljaoca (default: noreply@<tenant>.onmicrosoft.com).

.PARAMETER SmtpServer
    SMTP server za slanje email-a (default: smtp.office365.com).

.PARAMETER IncludeOnlyReady
    Ukljucuje u izvjestaj samo mailboxeve spremne za finalizaciju (>=95%).

.EXAMPLE
    .\Send-MigrationStatusReport.ps1 -To "admin@contoso.com"

    Salje izvjestaj o svim migracijama na navedenu adresu.

.EXAMPLE
    .\Send-MigrationStatusReport.ps1 -BatchName "Batch-Finance" -To "it-team@contoso.com" -IncludeOnlyReady

    Salje izvjestaj samo o mailboxevima spremnim za finalizaciju.

.NOTES
    Autor: PowerShell Migration Script
    Verzija: 2.0
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

function Format-DurationShort {
    param([timespan]$Duration)
    if ($Duration.TotalDays -ge 1) {
        return "{0:N0}d {1:N0}h {2:N0}m" -f [Math]::Floor($Duration.TotalDays), $Duration.Hours, $Duration.Minutes
    }
    elseif ($Duration.TotalHours -ge 1) {
        return "{0:N0}h {1:N0}m" -f [Math]::Floor($Duration.TotalHours), $Duration.Minutes
    }
    else {
        return "{0:N0}m" -f [Math]::Floor($Duration.TotalMinutes)
    }
}

function New-MigrationHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [array]$MigrationData,

        [Parameter(Mandatory = $false)]
        [switch]$OnlyReady
    )

    $timestamp = Get-Date -Format "dd.MM.yyyy HH:mm:ss"

    if ($OnlyReady) {
        $MigrationData = $MigrationData | Where-Object { $_.PercentageComplete -ge 95 }
    }

    $batchGroups = $MigrationData | Group-Object -Property BatchId

    $readyCount = ($MigrationData | Where-Object { $_.PercentageComplete -ge 95 }).Count
    $inProgressCount = ($MigrationData | Where-Object { $_.PercentageComplete -lt 95 }).Count

    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body { font-family: Arial, sans-serif; background-color: #f5f5f5; margin: 20px; }
.container { max-width: 1200px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; }
h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
h2 { color: #106ebe; margin-top: 25px; }
.summary { background-color: #f0f8ff; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 5px solid #0078d4; }
table { width: 100%; border-collapse: collapse; margin: 20px 0; }
th { background-color: #0078d4; color: white; padding: 12px; text-align: left; }
td { padding: 10px; border-bottom: 1px solid #e1e1e1; }
tr:hover { background-color: #f5f5f5; }
.status-ready { color: #107c10; font-weight: bold; }
.status-progress { color: #0078d4; }
.status-failed { color: #d13438; font-weight: bold; }
.footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #e1e1e1; color: #666; font-size: 12px; text-align: center; }
</style>
</head>
<body>
<div class="container">
<h1>Exchange Online Mailbox Migration - Status Izvjestaj</h1>
<p><strong>Datum/Vrijeme:</strong> $timestamp</p>

<div class="summary">
<div><strong>Ukupno batch-eva:</strong> $($batchGroups.Count)</div>
<div><strong>Ukupno mailboxeva:</strong> $($MigrationData.Count)</div>
<div class="status-ready"><strong>[OK] Spremno za finalizaciju (>=95%):</strong> $readyCount</div>
<div class="status-progress"><strong>[...] U toku (&lt;95%):</strong> $inProgressCount</div>
</div>
"@

    foreach ($batchGroup in $batchGroups) {
        $batchName = $batchGroup.Name
        $mailboxes = $batchGroup.Group | Sort-Object -Property PercentageComplete -Descending

        $readyInBatch = ($mailboxes | Where-Object { $_.PercentageComplete -ge 95 }).Count

        $htmlBody += @"
<h2>Batch: $batchName</h2>
<p>Spremno za finalizaciju: <strong>$readyInBatch/$($mailboxes.Count)</strong></p>
<table>
<thead>
<tr>
<th>Mailbox</th>
<th>Progress</th>
<th>Trajanje</th>
<th>GB Transferred</th>
</tr>
</thead>
<tbody>
"@

        foreach ($mailbox in $mailboxes) {
            $statusClass = if ($mailbox.PercentageComplete -ge 95) { 'status-ready' } else { 'status-progress' }

            $duration = if ($mailbox.SyncDuration) {
                Format-DurationShort -Duration $mailbox.SyncDuration
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

            $htmlBody += @"
<tr>
<td><strong>$($mailbox.Identity)</strong></td>
<td class="$statusClass">$($mailbox.PercentageComplete)%</td>
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
<p>Automatski generiran izvjestaj - PowerShell Migration Scripts</p>
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
        $batches = @(Get-MigrationBatch -ErrorAction Stop | Where-Object { $_.Status.Value -ne 4 -and $_.Status -ne 'Completed' })
    }

    if ($batches.Count -eq 0) {
        Write-Warning "Nisu pronadeni migration batch-evi."
        return
    }

    # Prikupi sve podatke
    $allMigrationData = @()

    foreach ($batch in $batches) {
        $users = @(Get-MigrationUser -BatchId $batch.Identity -ErrorAction SilentlyContinue)
        $batchStartTime = $batch.CreationDateTime

        foreach ($user in $users) {
            $stats = Get-MigrationUserStatistics -Identity $user.Identity -ErrorAction SilentlyContinue

            if ($stats) {
                $syncDuration = $null
                $startTime = $null

                if ($stats.InitialSyncDateTime) {
                    $startTime = $stats.InitialSyncDateTime
                }
                elseif ($stats.StartDate) {
                    $startTime = $stats.StartDate
                }
                elseif ($batchStartTime) {
                    $startTime = $batchStartTime
                }

                if ($startTime) {
                    $endTime = if ($stats.LastSyncedDateTime) { $stats.LastSyncedDateTime } else { Get-Date }
                    $syncDuration = $endTime - $startTime
                }

                $allMigrationData += [PSCustomObject]@{
                    Identity           = $stats.Identity
                    BatchId            = $stats.BatchId
                    Status             = $stats.Status
                    PercentageComplete = if ($stats.PercentageComplete) { $stats.PercentageComplete } else { 0 }
                    BytesTransferred   = $stats.BytesTransferred
                    SyncDuration       = $syncDuration
                }
            }
        }
    }

    if ($allMigrationData.Count -eq 0) {
        Write-Warning "Nisu pronadeni migration users."
        return
    }

    Write-Host "Generisem HTML izvjestaj..." -ForegroundColor Cyan

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

    Write-Host "Saljem email izvjestaj na $($To -join ', ')..." -ForegroundColor Cyan

    try {
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

        Send-MailMessage @mailParams -ErrorAction Stop

        Write-Host "[OK] Email izvjestaj uspjesno poslan!" -ForegroundColor Green
    }
    catch {
        Write-Warning "Nije moguce poslati email putem SMTP: $_"
        Write-Host ""
        Write-Host "[i] Alternativa: Sacuvajte HTML izvjestaj lokalno:" -ForegroundColor Cyan

        $reportPath = ".\MigrationReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
        $htmlReport | Out-File -FilePath $reportPath -Encoding UTF8

        Write-Host "   Izvjestaj sacuvan u: $reportPath" -ForegroundColor Green
        Write-Host "   Mozete ga poslati rucno ili koristiti Graph API za slanje." -ForegroundColor Yellow
    }
}
catch {
    Write-Error "Greska: $_"
    Write-Error $_.Exception.Message
}
