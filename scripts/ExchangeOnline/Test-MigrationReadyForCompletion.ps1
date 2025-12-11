<#
.SYNOPSIS
    Brza provjera koji migration batch-evi i mailboxevi su spremni za finalizaciju (>=95%).

.DESCRIPTION
    Jednostavna skripta koja prikazuje samo one migration batch-eve i mailboxeve
    koji su dostigli 95% ili vise i spremni su za pokretanje Complete-MigrationBatch komande.

.PARAMETER BatchName
    Naziv specificnog migration batcha za provjeru.

.PARAMETER AutoComplete
    Automatski pokrece Complete-MigrationBatch za sve batch-eve spremne za finalizaciju.
    OPREZ: Ovo ce pokrenuti finalizaciju bez dodatne potvrde!

.EXAMPLE
    .\Test-MigrationReadyForCompletion.ps1

    Prikazuje sve migration batch-eve i mailboxeve spremne za finalizaciju.

.EXAMPLE
    .\Test-MigrationReadyForCompletion.ps1 -BatchName "Batch-Finance-2024"

    Provjerava da li je specificni batch spreman za finalizaciju.

.EXAMPLE
    .\Test-MigrationReadyForCompletion.ps1 -AutoComplete

    Automatski finalizira sve batch-eve spremne za completion (OPREZ!).

.NOTES
    Autor: PowerShell Migration Script
    Verzija: 2.0
    Datum: 2025-12-11
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$BatchName,

    [Parameter(Mandatory = $false)]
    [switch]$AutoComplete
)

#Requires -Modules ExchangeOnlineManagement

# Funkcija za formatiranje trajanja
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

try {
    # Provjera konekcije
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
    }
    catch {
        Write-Error "Nije uspostavljena konekcija sa Exchange Online. Pokrenite 'Connect-ExchangeOnline'."
        return
    }

    Write-Host ""
    Write-Host "===============================================================" -ForegroundColor Green
    Write-Host "  Provjera spremnosti za finalizaciju migracije (>=95%)     " -ForegroundColor Green
    Write-Host "===============================================================" -ForegroundColor Green
    Write-Host ""

    # Dohvati migration batch-eve
    if ($BatchName) {
        $batches = @(Get-MigrationBatch -Identity $BatchName -ErrorAction Stop)
    }
    else {
        $batches = @(Get-MigrationBatch -ErrorAction Stop | Where-Object { $_.Status.Value -ne 4 -and $_.Status -ne 'Completed' })
    }

    if ($batches.Count -eq 0) {
        Write-Host "[!] Nisu pronadeni migration batch-evi." -ForegroundColor Yellow
        return
    }

    $readyBatches = @()
    $readyMailboxes = @()

    foreach ($batch in $batches) {
        # Dohvati migration users
        $users = @(Get-MigrationUser -BatchId $batch.Identity -ErrorAction SilentlyContinue)

        if ($users.Count -eq 0) {
            continue
        }

        $usersReadyForCompletion = @()
        $batchStartTime = $batch.CreationDateTime

        foreach ($user in $users) {
            $stats = Get-MigrationUserStatistics -Identity $user.Identity -ErrorAction SilentlyContinue

            if ($stats -and $stats.PercentageComplete -ge 95) {
                # Izracunaj trajanje
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

                $usersReadyForCompletion += [PSCustomObject]@{
                    Identity           = $stats.Identity
                    BatchId            = $stats.BatchId
                    PercentageComplete = $stats.PercentageComplete
                    Status             = $stats.Status
                    BytesTransferred   = $stats.BytesTransferred
                    TotalMailboxSize   = $stats.TotalMailboxSize
                    SyncDuration       = $syncDuration
                }

                $readyMailboxes += $stats.Identity
            }
        }

        # Ako svi mailboxevi u batchu su >=95%, batch je spreman
        if ($usersReadyForCompletion.Count -eq $users.Count -and $usersReadyForCompletion.Count -gt 0) {
            $readyBatches += $batch.Identity

            Write-Host "[OK] BATCH SPREMAN: " -NoNewline -ForegroundColor Green
            Write-Host $batch.Identity -ForegroundColor White
            Write-Host "   Mailboxeva: $($usersReadyForCompletion.Count)/$($users.Count) (100%)" -ForegroundColor Green
            Write-Host "   Mailboxevi spremni za finalizaciju:" -ForegroundColor Cyan

            foreach ($readyUser in $usersReadyForCompletion) {
                $duration = if ($readyUser.SyncDuration) {
                    Format-DurationShort -Duration $readyUser.SyncDuration
                }
                else {
                    "N/A"
                }

                Write-Host "      * $($readyUser.Identity) - $($readyUser.PercentageComplete)% (trajanje: $duration)" -ForegroundColor White
            }

            Write-Host ""
            Write-Host "   [>] Komanda za finalizaciju:" -ForegroundColor Yellow
            Write-Host "      Complete-MigrationBatch -Identity '$($batch.Identity)'" -ForegroundColor White
            Write-Host ""
        }
        elseif ($usersReadyForCompletion.Count -gt 0) {
            Write-Host "[!] BATCH DJELOMICNO SPREMAN: " -NoNewline -ForegroundColor Yellow
            Write-Host $batch.Identity -ForegroundColor White
            Write-Host "   Mailboxeva spremno: $($usersReadyForCompletion.Count)/$($users.Count) ($([Math]::Round(($usersReadyForCompletion.Count / $users.Count) * 100, 1))%)" -ForegroundColor Yellow
            Write-Host "   Spremni mailboxevi:" -ForegroundColor Cyan

            foreach ($readyUser in $usersReadyForCompletion) {
                $duration = if ($readyUser.SyncDuration) {
                    Format-DurationShort -Duration $readyUser.SyncDuration
                }
                else {
                    "N/A"
                }

                Write-Host "      * $($readyUser.Identity) - $($readyUser.PercentageComplete)% (trajanje: $duration)" -ForegroundColor White
            }

            Write-Host ""
            Write-Host "   [!] Preostali mailboxevi jos nisu dostigli 95%" -ForegroundColor Yellow
            Write-Host "   [i] Pricekajte da svi mailboxevi dostignu 95% prije finalizacije batcha" -ForegroundColor Cyan
            Write-Host ""
        }
        else {
            Write-Host "[...] Batch u toku: " -NoNewline -ForegroundColor Cyan
            Write-Host $batch.Identity -ForegroundColor White
            Write-Host "   Nijedan mailbox jos nije dostigao 95%" -ForegroundColor Cyan
            Write-Host ""
        }
    }

    # Sazetak
    Write-Host "===============================================================" -ForegroundColor Green
    Write-Host "  Sazetak                                                    " -ForegroundColor Green
    Write-Host "===============================================================" -ForegroundColor Green

    Write-Host ""
    Write-Host "  Batch-eva provjereno: $($batches.Count)"
    Write-Host "  Batch-eva spremno za finalizaciju: " -NoNewline
    Write-Host $readyBatches.Count -ForegroundColor $(if ($readyBatches.Count -gt 0) { 'Green' } else { 'Yellow' })

    Write-Host "  Mailboxeva spremno za finalizaciju: " -NoNewline
    Write-Host $readyMailboxes.Count -ForegroundColor $(if ($readyMailboxes.Count -gt 0) { 'Green' } else { 'Yellow' })
    Write-Host ""

    # Auto-complete ako je trazeno
    if ($AutoComplete -and $readyBatches.Count -gt 0) {
        Write-Host "[!!!] UPOZORENJE: AutoComplete je omogucen!" -ForegroundColor Red
        Write-Host "   Pokrecem finalizaciju za $($readyBatches.Count) batch(eva)..." -ForegroundColor Red
        Write-Host ""

        foreach ($batchToComplete in $readyBatches) {
            Write-Host "   [>] Finaliziram: $batchToComplete..." -ForegroundColor Yellow

            try {
                Complete-MigrationBatch -Identity $batchToComplete -Confirm:$false -ErrorAction Stop
                Write-Host "   [OK] Uspjesno pokrenuta finalizacija za: $batchToComplete" -ForegroundColor Green
            }
            catch {
                Write-Host "   [ERROR] Greska pri finalizaciji $batchToComplete : $_" -ForegroundColor Red
            }

            Write-Host ""
        }
    }
    elseif ($readyBatches.Count -gt 0) {
        Write-Host "[i] Savjet: Dodajte '-AutoComplete' parametar za automatsku finalizaciju" -ForegroundColor Cyan
        Write-Host "   (OPREZ: Finalizacija ce se pokrenuti bez dodatne potvrde!)" -ForegroundColor Yellow
        Write-Host ""
    }

    if ($readyBatches.Count -eq 0) {
        Write-Host "[...] Nijedan batch jos nije spreman za finalizaciju." -ForegroundColor Yellow
        Write-Host "   Pricekajte da mailboxevi dostignu 95% completion." -ForegroundColor Cyan
        Write-Host ""
    }

    Write-Host "===============================================================" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Error "Greska: $_"
    Write-Error $_.Exception.Message
}
