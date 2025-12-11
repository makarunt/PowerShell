<#
.SYNOPSIS
    Izvlaci trajanje remote mailbox migracija sa on-premises Exchange servera na Exchange Online.

.DESCRIPTION
    Ova skripta omogucava jednostavno pracenje vremena trajanja mailbox migracija do 95% completion,
    sto je faza kada se moze pokrenuti finalizacija (Complete-MigrationBatch).

    Skripta dohvaca detaljne statistike za migration batch(eve) i prikazuje:
    - Vrijeme trajanja do 95% completion
    - Trenutni status migracije
    - Detaljne statistike po mailboxu
    - Procjenu preostalog vremena (ako je migracija u toku)

.PARAMETER BatchName
    Naziv specificnog migration batcha. Ako nije naveden, prikazuju se svi batchevi.

.PARAMETER IncludeCompleted
    Ukljucuje completed migration batch-eve u rezultate.

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

    Prikazuje detaljne statistike za specificni migration batch.

.EXAMPLE
    .\Get-MailboxMigrationDuration.ps1 -IncludeCompleted -ExportToCsv

    Izvozi sve migracije (ukljucujuci completed) u CSV datoteku.

.NOTES
    Autor: PowerShell Migration Script
    Verzija: 2.0
    Datum: 2025-12-11

    Zahtjevi:
    - ExchangeOnlineManagement modul
    - Aktivna konekcija na Exchange Online (Connect-ExchangeOnline)
    - Odgovarajuce permisije za citanje migration podataka

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
        return "{0:N0} dana, {1:N0} sati, {2:N0} minuta" -f [Math]::Floor($Duration.TotalDays), $Duration.Hours, $Duration.Minutes
    }
    elseif ($Duration.TotalHours -ge 1) {
        return "{0:N0} sati, {1:N0} minuta" -f [Math]::Floor($Duration.TotalHours), $Duration.Minutes
    }
    else {
        return "{0:N0} minuta" -f [Math]::Floor($Duration.TotalMinutes)
    }
}

# Funkcija za izracun vremena do 95%
function Get-MigrationDurationInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [object]$MigrationUser,

        [Parameter(Mandatory = $false)]
        [datetime]$BatchCreationTime,

        [Parameter(Mandatory = $false)]
        [object]$BatchStatus,

        [Parameter(Mandatory = $false)]
        [datetime]$BatchLastSyncedDateTime
    )

    # Ako je batch completed, koristi batch status umjesto user statusa (koji moze biti pogresan)
    $effectiveStatus = $MigrationUser.Status
    if ($BatchStatus) {
        $batchStatusText = Get-StatusText -Status $BatchStatus
        if ($batchStatusText -eq 'Completed' -and $MigrationUser.PercentageComplete -eq 100) {
            # Override: batch je completed i mailbox je 100%, koristi batch status
            $effectiveStatus = $BatchStatus
        }
    }

    $result = [PSCustomObject]@{
        Identity                = $MigrationUser.Identity
        BatchId                 = $MigrationUser.BatchId
        Status                  = $effectiveStatus
        PercentageComplete      = if ($MigrationUser.PercentageComplete) { $MigrationUser.PercentageComplete } else { 0 }
        StartDate               = $MigrationUser.StartDate
        InitialSyncDateTime     = $MigrationUser.InitialSyncDateTime
        LastSyncedDateTime      = $MigrationUser.LastSyncedDateTime
        CompletionDateTime      = $MigrationUser.CompletionDateTime
        TimeTo95Percent         = $null
        TimeTo95PercentFormatted = "N/A"
        TotalDuration           = $null
        TotalDurationFormatted  = "N/A"
        CurrentDuration         = $null
        CurrentDurationFormatted = "N/A"
        EstimatedTimeRemaining  = $null
        EstimatedCompletion     = $null
        BytesTransferred        = $MigrationUser.BytesTransferred
        TotalMailboxSize        = $MigrationUser.TotalMailboxSize
        ItemsTransferred        = $MigrationUser.ItemsTransferred
        TotalItemsInMailbox     = $MigrationUser.TotalItemsInMailbox
        IsReadyForFinalization  = $false
        Message                 = ""
    }

    # Odredi pocetno vrijeme - koristi prvi dostupan datum
    $startTime = $null
    if ($MigrationUser.InitialSyncDateTime) {
        $startTime = $MigrationUser.InitialSyncDateTime
    }
    elseif ($MigrationUser.StartDate) {
        $startTime = $MigrationUser.StartDate
    }
    elseif ($MigrationUser.QueuedDateTime) {
        $startTime = $MigrationUser.QueuedDateTime
    }
    elseif ($BatchCreationTime) {
        $startTime = $BatchCreationTime
    }

    # Ako nemamo pocetno vrijeme, probaj sa CreationDatetime
    if (-not $startTime -and $MigrationUser.CreationDateTime) {
        $startTime = $MigrationUser.CreationDateTime
    }

    # Ako jos uvijek nemamo podatke
    if (-not $startTime) {
        $result.Message = "Migracija jos nije zapocela - nema vremenskih podataka"
        return $result
    }

    # Odredi zavrsno vrijeme - VAZNO: za completed migracije koristiti stvarno vrijeme zavrsetka!
    $endTime = Get-Date

    # Za completed migracije, trazi razlicite properties koji oznacavaju zavrsno vrijeme
    # Provjeri i effectiveStatus (koji ukljucuje batch status override)
    $isCompleted = ($MigrationUser.PercentageComplete -eq 100) -or
                   ($effectiveStatus -eq 'Completed') -or
                   ($effectiveStatus -eq 4) -or
                   ($effectiveStatus.Value -eq 4) -or
                   ($MigrationUser.Status -eq 'Completed') -or
                   ($MigrationUser.Status -eq 4) -or
                   ($MigrationUser.Status.Value -eq 4)

    if ($isCompleted) {
        # Za zavrsene migracije, pokusaj razlicite sources za completion time
        # PRIORITET: BatchLastSyncedDateTime (ako je batch completed)
        if ($BatchLastSyncedDateTime) {
            $endTime = $BatchLastSyncedDateTime
        }
        elseif ($MigrationUser.CompletionDateTime) {
            $endTime = $MigrationUser.CompletionDateTime
        }
        elseif ($MigrationUser.FinalizationDateTime) {
            $endTime = $MigrationUser.FinalizationDateTime
        }
        elseif ($MigrationUser.LastSuccessfulSyncTime) {
            $endTime = $MigrationUser.LastSuccessfulSyncTime
        }
        elseif ($MigrationUser.LastSyncedDateTime) {
            $endTime = $MigrationUser.LastSyncedDateTime
        }
        else {
            # Ako nista drugo, koristi Report ako postoji
            if ($MigrationUser.Report -and $MigrationUser.Report.Entries) {
                $lastEntry = $MigrationUser.Report.Entries | Sort-Object Date -Descending | Select-Object -First 1
                if ($lastEntry -and $lastEntry.Date) {
                    $endTime = $lastEntry.Date
                }
            }
        }
    }
    else {
        # Za migracije u toku, koristi Last synced ili sada
        if ($MigrationUser.LastSyncedDateTime) {
            $endTime = $MigrationUser.LastSyncedDateTime
        }
    }

    # Izracunaj trenutno trajanje
    $currentDuration = $endTime - $startTime
    $result.CurrentDuration = $currentDuration
    $result.CurrentDurationFormatted = Format-Duration -Duration $currentDuration

    # Koristi vec setovanu $isCompleted varijablu (setovana gore)
    # Ako je migracija zavrsena (100% ili Completed status)
    if ($isCompleted) {
        $result.TotalDuration = $currentDuration
        $result.TotalDurationFormatted = Format-Duration -Duration $currentDuration

        # Izracunaj vrijeme do 95% - TimeSpan ne moze se mnoziti direktno!
        $minutesTo95 = $currentDuration.TotalMinutes * (95.0 / 100.0)
        $result.TimeTo95Percent = [timespan]::FromMinutes($minutesTo95)
        $result.TimeTo95PercentFormatted = Format-Duration -Duration $result.TimeTo95Percent

        $result.IsReadyForFinalization = $true
        $result.Message = "[OK] Zavrseno - ukupno trajanje"
        return $result
    }

    # Provjeri da li je dosegnuto 95%
    if ($MigrationUser.PercentageComplete -ge 95) {
        $result.IsReadyForFinalization = $true

        # Izracunaj vrijeme do 95% na osnovu trenutnog postotka
        $currentPercent = [Math]::Max($MigrationUser.PercentageComplete, 1)

        # TimeSpan ne moze se mnoziti direktno - koristimo minute!
        $minutesTo95 = $currentDuration.TotalMinutes * (95.0 / $currentPercent)
        $timeTo95 = [timespan]::FromMinutes($minutesTo95)

        $result.TimeTo95Percent = $timeTo95
        $result.TimeTo95PercentFormatted = Format-Duration -Duration $timeTo95

        $result.TotalDuration = $currentDuration
        $result.TotalDurationFormatted = Format-Duration -Duration $currentDuration

        $result.Message = "[OK] SPREMAN ZA FINALIZACIJU - Complete-MigrationBatch"
    }
    else {
        # Procijeni preostalo vrijeme
        $currentPercent = [Math]::Max($MigrationUser.PercentageComplete, 1)

        if ($currentPercent -gt 0 -and $currentDuration.TotalMinutes -gt 0) {
            $timePerPercent = $currentDuration.TotalMinutes / $currentPercent
            $remainingPercent = 95 - $currentPercent
            $estimatedMinutesTo95 = $timePerPercent * $remainingPercent

            if ($estimatedMinutesTo95 -gt 0) {
                $estimatedTimespan = [timespan]::FromMinutes($estimatedMinutesTo95)
                $result.EstimatedTimeRemaining = $estimatedTimespan
                $result.EstimatedCompletion = (Get-Date).Add($estimatedTimespan)
                $result.Message = "U toku - procjena do 95%: " + (Format-Duration -Duration $estimatedTimespan)
            }
            else {
                $result.Message = "U toku - migracija u tijeku"
            }
        }
        else {
            $result.Message = "U toku - pocetna faza"
        }

        $result.TotalDuration = $currentDuration
        $result.TotalDurationFormatted = Format-Duration -Duration $currentDuration
    }

    return $result
}

# Funkcija za konverziju statusa u tekst
function Get-StatusText {
    param([object]$Status)

    # Ako je status broj, mapiranje na tekstualne vrijednosti
    $statusMap = @{
        0 = 'Created'
        1 = 'Syncing'
        2 = 'Synced'
        3 = 'Failed'
        4 = 'Completed'
        5 = 'Stopped'
        6 = 'Queued'
    }

    # Pokusaj izvuci integer vrijednost iz razlicitih formata
    $statusValue = $null

    if ($Status -is [int]) {
        $statusValue = $Status
    }
    elseif ($Status.Value -ne $null) {
        # Status je objekt sa .Value property
        $statusValue = $Status.Value
    }
    elseif ($Status -is [string]) {
        # Ako je vec string, vrati ga
        return $Status
    }

    # Ako imamo integer vrijednost, mapiraj ga
    if ($statusValue -ne $null -and $statusMap.ContainsKey($statusValue)) {
        return $statusMap[$statusValue]
    }

    # Fallback - vrati original
    return $Status.ToString()
}

# Glavna skripta
try {
    Write-Verbose "Provjeravam Exchange Online konekciju..."

    if (-not (Test-ExchangeOnlineConnection)) {
        Write-Error "Nije uspostavljena konekcija sa Exchange Online. Pokrenite 'Connect-ExchangeOnline' prije izvrsavanja ove skripte."
        return
    }

    Write-Host ""
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host "   Analiza trajanja mailbox migracija na Exchange Online" -ForegroundColor Cyan
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host ""

    # Dohvati migration batch-eve
    Write-Verbose "Dohvacam migration batch-eve..."

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

    # Filtriraj completed batch-eve ako nije ukljuceno
    if (-not $IncludeCompleted) {
        $batches = $batches | Where-Object { $_.Status.Value -ne 4 -and $_.Status -ne 'Completed' }
    }

    if ($batches.Count -eq 0) {
        Write-Warning "Nisu pronadeni migration batch-evi koji zadovoljavaju kriterije."
        return
    }

    Write-Host "Pronadeno $($batches.Count) migration batch(eva)" -ForegroundColor Green
    Write-Host ""

    # Obrada svakog batch-a
    $allResults = @()

    foreach ($batch in $batches) {
        Write-Host "---------------------------------------------------------------" -ForegroundColor Yellow
        Write-Host "Batch: $($batch.Identity)" -ForegroundColor Yellow
        Write-Host "---------------------------------------------------------------" -ForegroundColor Yellow

        $batchStatusText = Get-StatusText -Status $batch.Status
        Write-Host "  Status: $batchStatusText"
        Write-Host "  Total Mailboxes: $($batch.TotalCount)"
        Write-Host "  Synced: $($batch.SyncedItemCount)"
        Write-Host "  Active: $($batch.ActiveItemCount)"
        Write-Host "  Failed: $($batch.FailedItemCount)"

        if ($batch.CreationDateTime) {
            Write-Host "  Created: $($batch.CreationDateTime.ToString('dd.MM.yyyy HH:mm:ss'))"

            # Izracunaj batch trajanje - za completed batch koristi LastSyncedDateTime!
            $batchStatusText = Get-StatusText -Status $batch.Status
            $batchIsCompleted = ($batchStatusText -eq 'Completed') -or ($batch.Status -eq 'Completed') -or ($batch.Status -eq 4) -or ($batch.Status.Value -eq 4)

            $batchEndTime = Get-Date
            if ($batchIsCompleted -and $batch.LastSyncedDateTime) {
                $batchEndTime = $batch.LastSyncedDateTime
            }
            elseif ($batch.LastSyncedDateTime) {
                $batchEndTime = $batch.LastSyncedDateTime
            }

            $batchDuration = $batchEndTime - $batch.CreationDateTime
            Write-Host "  Batch traje: $(Format-Duration -Duration $batchDuration)" -ForegroundColor Cyan
        }

        Write-Host ""

        # Dohvati migration users za ovaj batch
        Write-Verbose "Dohvacam migration users za batch '$($batch.Identity)'..."

        $migrationUsers = @(Get-MigrationUser -BatchId $batch.Identity -ErrorAction Stop)

        if ($migrationUsers.Count -eq 0) {
            Write-Warning "  Nisu pronadeni migration users za ovaj batch."
            Write-Host ""
            continue
        }

        Write-Host "  Analiziram $($migrationUsers.Count) mailbox(eva)..." -ForegroundColor Cyan
        Write-Host ""

        # Obrada svakog migration user-a
        foreach ($user in $migrationUsers) {
            # Dohvati detaljne statistike
            $userStats = Get-MigrationUserStatistics -Identity $user.Identity -ErrorAction SilentlyContinue

            if (-not $userStats) {
                Write-Warning "  Nije moguce dohvatiti statistiku za $($user.Identity)"
                continue
            }

            # Izracunaj vrijeme do 95%
            # Proslijedi batch status i batch LastSyncedDateTime za tocnije rezultate
            $migrationData = Get-MigrationDurationInfo `
                -MigrationUser $userStats `
                -BatchCreationTime $batch.CreationDateTime `
                -BatchStatus $batch.Status `
                -BatchLastSyncedDateTime $batch.LastSyncedDateTime
            $allResults += $migrationData

            # Prikazuj osnovne informacije
            $userStatusText = Get-StatusText -Status $migrationData.Status
            $statusColor = switch ($userStatusText) {
                'Synced' { 'Green' }
                'Completed' { 'Green' }
                'Syncing' { 'Cyan' }
                'Failed' { 'Red' }
                'Queued' { 'Yellow' }
                default { 'White' }
            }

            Write-Host "  [>] $($migrationData.Identity)" -ForegroundColor White
            Write-Host "      Status: " -NoNewline
            Write-Host $userStatusText -ForegroundColor $statusColor -NoNewline
            Write-Host " | Progress: $($migrationData.PercentageComplete)%"

            # VAZNO: Uvijek prikazi trenutno trajanje
            if ($migrationData.CurrentDurationFormatted -ne "N/A") {
                Write-Host "      TRAJANJE: " -NoNewline -ForegroundColor White
                Write-Host $migrationData.CurrentDurationFormatted -ForegroundColor Cyan
            }

            if ($migrationData.TimeTo95PercentFormatted -ne "N/A" -and $migrationData.PercentageComplete -ge 95) {
                Write-Host "      Vrijeme do 95%: " -NoNewline -ForegroundColor Green
                Write-Host $migrationData.TimeTo95PercentFormatted -ForegroundColor Green
            }

            if ($migrationData.IsReadyForFinalization) {
                Write-Host "      >>> " -NoNewline -ForegroundColor Green
                Write-Host "SPREMAN ZA FINALIZACIJU!" -ForegroundColor Green
            }
            elseif ($migrationData.EstimatedTimeRemaining) {
                $estimatedText = Format-Duration -Duration $migrationData.EstimatedTimeRemaining
                Write-Host "      Procjena do 95%: " -NoNewline
                Write-Host $estimatedText -ForegroundColor Yellow
                if ($migrationData.EstimatedCompletion) {
                    Write-Host "      Procijenjeno vrijeme: $($migrationData.EstimatedCompletion.ToString('dd.MM.yyyy HH:mm'))" -ForegroundColor Yellow
                }
            }

            # Detaljne statistike ako je trazeno
            if ($ShowDetailedStats) {
                Write-Host "      ---Detaljne statistike---" -ForegroundColor DarkGray

                if ($migrationData.BytesTransferred) {
                    $bytesGB = [Math]::Round($migrationData.BytesTransferred / 1GB, 2)
                    Write-Host "      Bytes Transferred: $bytesGB GB"
                }

                if ($migrationData.TotalMailboxSize) {
                    $sizeGB = [Math]::Round($migrationData.TotalMailboxSize / 1GB, 2)
                    Write-Host "      Total Mailbox Size: $sizeGB GB"
                }

                Write-Host "      Items Transferred: $($migrationData.ItemsTransferred)"
                Write-Host "      Total Items: $($migrationData.TotalItemsInMailbox)"

                if ($migrationData.InitialSyncDateTime) {
                    Write-Host "      Sync Started: $($migrationData.InitialSyncDateTime.ToString('dd.MM.yyyy HH:mm:ss'))"
                }

                if ($migrationData.LastSyncedDateTime) {
                    Write-Host "      Last Synced: $($migrationData.LastSyncedDateTime.ToString('dd.MM.yyyy HH:mm:ss'))"
                }
            }

            Write-Host ""
        }

        Write-Host ""
    }

    # Sazetak
    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host "   Sazetak" -ForegroundColor Cyan
    Write-Host "===============================================================" -ForegroundColor Cyan

    $readyCount = ($allResults | Where-Object { $_.IsReadyForFinalization }).Count
    $inProgressCount = ($allResults | Where-Object { $_.PercentageComplete -lt 95 -and ($_.Status -eq 'Syncing' -or $_.Status -eq 1) }).Count
    $failedCount = ($allResults | Where-Object { $_.Status -eq 'Failed' -or $_.Status -eq 3 }).Count

    Write-Host "  Ukupno mailboxeva: $($allResults.Count)"
    Write-Host "  Spremno za finalizaciju (>=95%): " -NoNewline
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
        Write-Host "[OK] Mailboxevi spremni za finalizaciju (pokrenite Complete-MigrationBatch):" -ForegroundColor Green
        $allResults | Where-Object { $_.IsReadyForFinalization } | ForEach-Object {
            Write-Host "  * $($_.Identity) - dosegnuto $($_.PercentageComplete)% za $($_.TimeTo95PercentFormatted)" -ForegroundColor Green
        }
        Write-Host ""
    }

    # Export u CSV ako je trazeno
    if ($ExportToCsv) {
        if (-not $CsvPath) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $CsvPath = ".\MigrationDuration_$timestamp.csv"
        }

        Write-Host "Izvozim rezultate u CSV: $CsvPath" -ForegroundColor Cyan

        $allResults | Select-Object Identity, BatchId, Status, PercentageComplete,
                                    IsReadyForFinalization, TimeTo95PercentFormatted,
                                    CurrentDurationFormatted, TotalDurationFormatted,
                                    StartDate, InitialSyncDateTime,
                                    LastSyncedDateTime, EstimatedCompletion, BytesTransferred,
                                    TotalMailboxSize, ItemsTransferred, TotalItemsInMailbox,
                                    Message |
            Export-Csv -Path $CsvPath -NoTypeInformation -Encoding UTF8

        Write-Host "[OK] CSV export zavrsen!" -ForegroundColor Green
        Write-Host ""
    }

    Write-Host "===============================================================" -ForegroundColor Cyan
    Write-Host ""
}
catch {
    Write-Error "Greska prilikom izvrsavanja skripte: $_"
    Write-Error $_.Exception.Message

    if ($_.Exception.InnerException) {
        Write-Error "Inner Exception: $($_.Exception.InnerException.Message)"
    }
}
