<#
.SYNOPSIS
    Izvlači trajanje remote mailbox migracija sa on-premises Exchange servera na Exchange Online.

.DESCRIPTION
    Ova skripta omogućava jednostavno praćenje vremena trajanja mailbox migracija do 95% completion,
    što je faza kada se može pokrenuti finalizacija (Complete-MigrationBatch).

    Skripta dohvaća detaljne statistike za migration batch(eve) i prikazuje:
    - Vrijeme trajanja do 95% completion
    - Trenutni status migracije
    - Detaljne statistike po mailboxu
    - Procjenu preostalog vremena (ako je migracija u toku)

.PARAMETER BatchName
    Naziv specifičnog migration batcha. Ako nije naveden, prikazuju se svi batchevi.

.PARAMETER IncludeCompleted
    Uključuje completed migration batch-eve u rezultate.

.PARAMETER ExportToCsv
    Izvozi rezultate u CSV datoteku.

.PARAMETER CsvPath
    Putanja za CSV export. Default je trenutni direktorij sa timestamp-om.

.PARAMETER ShowDetailedStats
    Prikazuje detaljne statistike za svaki mailbox u batchu.

.EXAMPLE
    .\Get-MailboxMigrationDuration.ps1

    Prikazuje trajanje migracija za sve aktivne migration batch-eve.

.EXAMPLE
    .\Get-MailboxMigrationDuration.ps1 -BatchName "Batch-Finance-2024" -ShowDetailedStats

    Prikazuje detaljne statistike za specifični migration batch.

.EXAMPLE
    .\Get-MailboxMigrationDuration.ps1 -IncludeCompleted -ExportToCsv

    Izvozi sve migracije (uključujući completed) u CSV datoteku.

.NOTES
    Autor: PowerShell Migration Script
    Verzija: 1.0
    Datum: 2025-12-11

    Zahtjevi:
    - ExchangeOnlineManagement modul
    - Aktivna konekcija na Exchange Online (Connect-ExchangeOnline)
    - Odgovarajuće permisije za čitanje migration podataka

.LINK
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-migrationbatch
    https://docs.microsoft.com/en-us/powershell/module/exchange/get-migrationuser
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false, Position = 0)]
    [ValidateNotNullOrEmpty()]
    [string]$BatchName,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeCompleted,

    [Parameter(Mandatory = $false)]
    [switch]$ExportToCsv,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$CsvPath,

    [Parameter(Mandatory = $false)]
    [switch]$ShowDetailedStats
)

#Requires -Modules ExchangeOnlineManagement

# Funkcija za provjeru Exchange Online konekcije
function Test-ExchangeOnlineConnection {
    [CmdletBinding()]
    param()

    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Funkcija za formatiranje vremena trajanja
function Format-Duration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [timespan]$Duration
    )

    if ($Duration.TotalDays -ge 1) {
        return "{0:N1} dana, {1:N0} sati, {2:N0} minuta" -f $Duration.TotalDays, $Duration.Hours, $Duration.Minutes
    }
    elseif ($Duration.TotalHours -ge 1) {
        return "{0:N0} sati, {1:N0} minuta" -f $Duration.TotalHours, $Duration.Minutes
    }
    else {
        return "{0:N0} minuta" -f $Duration.TotalMinutes
    }
}

# Funkcija za izračun vremena do 95%
function Get-TimeTo95Percent {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$MigrationUser
    )

    $result = [PSCustomObject]@{
        Identity                = $MigrationUser.Identity
        BatchId                 = $MigrationUser.BatchId
        Status                  = $MigrationUser.Status
        PercentageComplete      = $MigrationUser.PercentageComplete
        StartDate               = $MigrationUser.StartDate
        InitialSyncDateTime     = $MigrationUser.InitialSyncDateTime
        LastSyncedDateTime      = $MigrationUser.LastSyncedDateTime
        CompletionDateTime      = $null
        TimeTo95Percent         = $null
        TimeTo95PercentFormatted = "N/A"
        TotalDuration           = $null
        TotalDurationFormatted  = "N/A"
        EstimatedTimeRemaining  = $null
        EstimatedCompletion     = $null
        BytesTransferred        = $MigrationUser.BytesTransferred
        TotalMailboxSize        = $MigrationUser.TotalMailboxSize
        ItemsTransferred        = $MigrationUser.ItemsTransferred
        TotalItemsInMailbox     = $MigrationUser.TotalItemsInMailbox
        IsReadyForFinalization  = $false
        Message                 = ""
    }

    # Ako nema podataka o sync-u, vrati prazan rezultat
    if (-not $MigrationUser.InitialSyncDateTime) {
        $result.Message = "Migracija još nije započela"
        return $result
    }

    # Izračunaj trenutno ukupno trajanje
    $endTime = $MigrationUser.LastSyncedDateTime
    if (-not $endTime) {
        $endTime = Get-Date
    }

    $totalDuration = $endTime - $MigrationUser.InitialSyncDateTime
    $result.TotalDuration = $totalDuration
    $result.TotalDurationFormatted = Format-Duration -Duration $totalDuration

    # Provjeri da li je dosegnuto 95%
    if ($MigrationUser.PercentageComplete -ge 95) {
        $result.IsReadyForFinalization = $true

        # Izračunaj vrijeme do 95% na osnovu trenutnog postotka
        # Pretpostavljamo linearnu progresiju za aproksimaciju
        $timeTo95 = $totalDuration * (95 / [Math]::Max($MigrationUser.PercentageComplete, 1))
        $result.TimeTo95Percent = $timeTo95
        $result.TimeTo95PercentFormatted = Format-Duration -Duration $timeTo95

        $result.CompletionDateTime = $MigrationUser.InitialSyncDateTime.Add($timeTo95)
        $result.Message = "✓ Spreman za finalizaciju (Complete-MigrationBatch)"
    }
    else {
        # Procijeni preostalo vrijeme
        $currentPercent = [Math]::Max($MigrationUser.PercentageComplete, 1)
        $timePerPercent = $totalDuration.TotalMinutes / $currentPercent
        $remainingPercent = 95 - $currentPercent
        $estimatedMinutesTo95 = $timePerPercent * $remainingPercent

        if ($estimatedMinutesTo95 -gt 0) {
            $estimatedTimespan = [timespan]::FromMinutes($estimatedMinutesTo95)
            $result.EstimatedTimeRemaining = $estimatedTimespan
            $result.EstimatedCompletion = (Get-Date).Add($estimatedTimespan)
            $result.Message = "U toku - procjena do 95%: " + (Format-Duration -Duration $estimatedTimespan)
        }
        else {
            $result.Message = "U toku - izračun procjene u tijeku..."
        }
    }

    return $result
}

# Glavna skripta
try {
    Write-Verbose "Provjeravam Exchange Online konekciju..."

    if (-not (Test-ExchangeOnlineConnection)) {
        Write-Error "Nije uspostavljena konekcija sa Exchange Online. Pokrenite 'Connect-ExchangeOnline' prije izvršavanja ove skripte."
        return
    }

    Write-Host "`n═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "   Analiza trajanja mailbox migracija na Exchange Online" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan

    # Dohvati migration batch-eve
    Write-Verbose "Dohvaćam migration batch-eve..."

    $batchParams = @{
        ErrorAction = 'Stop'
    }

    if ($BatchName) {
        $batchParams['Identity'] = $BatchName
        $batches = @(Get-MigrationBatch @batchParams)
    }
    else {
        $batches = @(Get-MigrationBatch @batchParams)
    }

    # Filtriraj completed batch-eve ako nije uključeno
    if (-not $IncludeCompleted) {
        $batches = $batches | Where-Object { $_.Status -ne 'Completed' }
    }

    if ($batches.Count -eq 0) {
        Write-Warning "Nisu pronađeni migration batch-evi koji zadovoljavaju kriterije."
        return
    }

    Write-Host "Pronađeno $($batches.Count) migration batch(eva)`n" -ForegroundColor Green

    # Obrada svakog batch-a
    $allResults = @()

    foreach ($batch in $batches) {
        Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor Yellow
        Write-Host "Batch: $($batch.Identity)" -ForegroundColor Yellow
        Write-Host "─────────────────────────────────────────────────────────────────" -ForegroundColor Yellow
        Write-Host "  Status: $($batch.Status)"
        Write-Host "  Total Mailboxes: $($batch.TotalCount)"
        Write-Host "  Synced: $($batch.SyncedItemCount)"
        Write-Host "  Active: $($batch.ActiveItemCount)"
        Write-Host "  Failed: $($batch.FailedItemCount)"

        if ($batch.CreationDateTime) {
            Write-Host "  Created: $($batch.CreationDateTime.ToString('dd.MM.yyyy HH:mm:ss'))"
        }

        Write-Host ""

        # Dohvati migration users za ovaj batch
        Write-Verbose "Dohvaćam migration users za batch '$($batch.Identity)'..."

        $migrationUsers = @(Get-MigrationUser -BatchId $batch.Identity -ErrorAction Stop)

        if ($migrationUsers.Count -eq 0) {
            Write-Warning "  Nisu pronađeni migration users za ovaj batch."
            continue
        }

        Write-Host "  Analiziram $($migrationUsers.Count) mailbox(eva)...`n" -ForegroundColor Cyan

        # Obrada svakog migration user-a
        foreach ($user in $migrationUsers) {
            # Dohvati detaljne statistike
            $userStats = Get-MigrationUserStatistics -Identity $user.Identity -ErrorAction SilentlyContinue

            if (-not $userStats) {
                Write-Warning "  Nije moguće dohvatiti statistiku za $($user.Identity)"
                continue
            }

            # Izračunaj vrijeme do 95%
            $migrationData = Get-TimeTo95Percent -MigrationUser $userStats
            $allResults += $migrationData

            # Prikaži osnovne informacije
            $statusColor = switch ($migrationData.Status) {
                'Synced' { 'Green' }
                'Syncing' { 'Cyan' }
                'Failed' { 'Red' }
                'Queued' { 'Yellow' }
                default { 'White' }
            }

            Write-Host "  📧 $($migrationData.Identity)" -ForegroundColor White
            Write-Host "     Status: " -NoNewline
            Write-Host $migrationData.Status -ForegroundColor $statusColor
            Write-Host "     Progress: $($migrationData.PercentageComplete)%"

            if ($migrationData.TimeTo95PercentFormatted -ne "N/A") {
                Write-Host "     ⏱️  Vrijeme do 95%: " -NoNewline -ForegroundColor Green
                Write-Host $migrationData.TimeTo95PercentFormatted -ForegroundColor Green
            }

            if ($migrationData.TotalDurationFormatted -ne "N/A") {
                Write-Host "     ⌚ Ukupno trajanje: $($migrationData.TotalDurationFormatted)"
            }

            if ($migrationData.IsReadyForFinalization) {
                Write-Host "     ✓ " -NoNewline -ForegroundColor Green
                Write-Host "SPREMAN ZA FINALIZACIJU!" -ForegroundColor Green
            }
            elseif ($migrationData.EstimatedTimeRemaining) {
                Write-Host "     ⏳ Procjena do 95%: $($migrationData.EstimatedCompletion.ToString('dd.MM.yyyy HH:mm'))"
            }

            # Detaljne statistike ako je traženo
            if ($ShowDetailedStats) {
                Write-Host "     ├─ Bytes Transferred: $($migrationData.BytesTransferred)"
                Write-Host "     ├─ Total Mailbox Size: $($migrationData.TotalMailboxSize)"
                Write-Host "     ├─ Items Transferred: $($migrationData.ItemsTransferred)"
                Write-Host "     ├─ Total Items: $($migrationData.TotalItemsInMailbox)"

                if ($migrationData.InitialSyncDateTime) {
                    Write-Host "     ├─ Sync Started: $($migrationData.InitialSyncDateTime.ToString('dd.MM.yyyy HH:mm:ss'))"
                }

                if ($migrationData.LastSyncedDateTime) {
                    Write-Host "     └─ Last Synced: $($migrationData.LastSyncedDateTime.ToString('dd.MM.yyyy HH:mm:ss'))"
                }
            }

            Write-Host ""
        }

        Write-Host ""
    }

    # Sažetak
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "   Sažetak" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════" -ForegroundColor Cyan

    $readyCount = ($allResults | Where-Object { $_.IsReadyForFinalization }).Count
    $inProgressCount = ($allResults | Where-Object { $_.PercentageComplete -lt 95 -and $_.Status -eq 'Syncing' }).Count
    $failedCount = ($allResults | Where-Object { $_.Status -eq 'Failed' }).Count

    Write-Host "  Ukupno mailboxeva: $($allResults.Count)"
    Write-Host "  Spremno za finalizaciju (≥95%): " -NoNewline
    Write-Host $readyCount -ForegroundColor Green
    Write-Host "  U toku (<95%): " -NoNewline
    Write-Host $inProgressCount -ForegroundColor Cyan

    if ($failedCount -gt 0) {
        Write-Host "  Neuspjelo: " -NoNewline
        Write-Host $failedCount -ForegroundColor Red
    }

    Write-Host ""

    # Prikaz mailboxeva spremnih za finalizaciju
    if ($readyCount -gt 0) {
        Write-Host "✓ Mailboxevi spremni za finalizaciju (pokrenite Complete-MigrationBatch):" -ForegroundColor Green
        $allResults | Where-Object { $_.IsReadyForFinalization } | ForEach-Object {
            Write-Host "  • $($_.Identity) - dosegnuto $($_.PercentageComplete)% za $($_.TimeTo95PercentFormatted)" -ForegroundColor Green
        }
        Write-Host ""
    }

    # Export u CSV ako je traženo
    if ($ExportToCsv) {
        if (-not $CsvPath) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $CsvPath = ".\MigrationDuration_$timestamp.csv"
        }

        Write-Host "Izvozim rezultate u CSV: $CsvPath" -ForegroundColor Cyan

        $allResults | Select-Object Identity, BatchId, Status, PercentageComplete,
                                    IsReadyForFinalization, TimeTo95PercentFormatted,
                                    TotalDurationFormatted, StartDate, InitialSyncDateTime,
                                    LastSyncedDateTime, EstimatedCompletion, BytesTransferred,
                                    TotalMailboxSize, ItemsTransferred, TotalItemsInMailbox,
                                    Message |
            Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

        Write-Host "✓ CSV export završen!" -ForegroundColor Green
        Write-Host ""
    }

    Write-Host "═══════════════════════════════════════════════════════════════`n" -ForegroundColor Cyan
}
catch {
    Write-Error "Greška prilikom izvršavanja skripte: $_"
    Write-Error $_.Exception.Message

    if ($_.Exception.InnerException) {
        Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
    }
}
