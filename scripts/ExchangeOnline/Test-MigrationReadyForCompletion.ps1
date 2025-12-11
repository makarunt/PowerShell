<#
.SYNOPSIS
    Brza provjera koji migration batch-evi i mailboxevi su spremni za finalizaciju (≥95%).

.DESCRIPTION
    Jednostavna skripta koja prikazuje samo one migration batch-eve i mailboxeve
    koji su dostigli 95% ili više i spremni su za pokretanje Complete-MigrationBatch komande.

.PARAMETER BatchName
    Naziv specifičnog migration batcha za provjeru.

.PARAMETER AutoComplete
    Automatski pokreće Complete-MigrationBatch za sve batch-eve spremne za finalizaciju.
    OPREZ: Ovo će pokrenuti finalizaciju bez dodatne potvrde!

.EXAMPLE
    .\Test-MigrationReadyForCompletion.ps1

    Prikazuje sve migration batch-eve i mailboxeve spremne za finalizaciju.

.EXAMPLE
    .\Test-MigrationReadyForCompletion.ps1 -BatchName "Batch-Finance-2024"

    Provjerava da li je specifični batch spreman za finalizaciju.

.EXAMPLE
    .\Test-MigrationReadyForCompletion.ps1 -AutoComplete

    Automatski finalizira sve batch-eve spremne za completion (OPREZ!).

.NOTES
    Autor: PowerShell Migration Script
    Verzija: 1.0
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

try {
    # Provjera konekcije
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
    }
    catch {
        Write-Error "Nije uspostavljena konekcija sa Exchange Online. Pokrenite 'Connect-ExchangeOnline'."
        return
    }

    Write-Host "`n╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║  Provjera spremnosti za finalizaciju migracije (≥95%)     ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════╝`n" -ForegroundColor Green

    # Dohvati migration batch-eve
    if ($BatchName) {
        $batches = @(Get-MigrationBatch -Identity $BatchName -ErrorAction Stop)
    }
    else {
        $batches = @(Get-MigrationBatch -ErrorAction Stop | Where-Object { $_.Status -ne 'Completed' })
    }

    if ($batches.Count -eq 0) {
        Write-Host "❌ Nisu pronađeni migration batch-evi." -ForegroundColor Yellow
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

        foreach ($user in $users) {
            $stats = Get-MigrationUserStatistics -Identity $user.Identity -ErrorAction SilentlyContinue

            if ($stats -and $stats.PercentageComplete -ge 95) {
                $usersReadyForCompletion += [PSCustomObject]@{
                    Identity           = $stats.Identity
                    BatchId            = $stats.BatchId
                    PercentageComplete = $stats.PercentageComplete
                    Status             = $stats.Status
                    BytesTransferred   = $stats.BytesTransferred
                    TotalMailboxSize   = $stats.TotalMailboxSize
                    SyncDuration       = if ($stats.InitialSyncDateTime -and $stats.LastSyncedDateTime) {
                        $stats.LastSyncedDateTime - $stats.InitialSyncDateTime
                    }
                    else { $null }
                }

                $readyMailboxes += $stats.Identity
            }
        }

        # Ako svi mailboxevi u batchu su ≥95%, batch je spreman
        if ($usersReadyForCompletion.Count -eq $users.Count -and $usersReadyForCompletion.Count -gt 0) {
            $readyBatches += $batch.Identity

            Write-Host "✅ BATCH SPREMAN: " -NoNewline -ForegroundColor Green
            Write-Host $batch.Identity -ForegroundColor White
            Write-Host "   📊 Mailboxeva: $($usersReadyForCompletion.Count)/$($users.Count) (100%)" -ForegroundColor Green
            Write-Host "   📧 Mailboxevi spremni za finalizaciju:" -ForegroundColor Cyan

            foreach ($readyUser in $usersReadyForCompletion) {
                $duration = if ($readyUser.SyncDuration) {
                    "{0:N0}d {1:N0}h {2:N0}m" -f $readyUser.SyncDuration.Days, $readyUser.SyncDuration.Hours, $readyUser.SyncDuration.Minutes
                }
                else {
                    "N/A"
                }

                Write-Host "      • $($readyUser.Identity) - $($readyUser.PercentageComplete)% (trajanje: $duration)" -ForegroundColor White
            }

            Write-Host "`n   ▶️  Komanda za finalizaciju:" -ForegroundColor Yellow
            Write-Host "      Complete-MigrationBatch -Identity '$($batch.Identity)'" -ForegroundColor White
            Write-Host ""
        }
        elseif ($usersReadyForCompletion.Count -gt 0) {
            Write-Host "⚠️  BATCH DJELOMIČNO SPREMAN: " -NoNewline -ForegroundColor Yellow
            Write-Host $batch.Identity -ForegroundColor White
            Write-Host "   📊 Mailboxeva spremno: $($usersReadyForCompletion.Count)/$($users.Count) ($([Math]::Round(($usersReadyForCompletion.Count / $users.Count) * 100, 1))%)" -ForegroundColor Yellow
            Write-Host "   📧 Spremni mailboxevi:" -ForegroundColor Cyan

            foreach ($readyUser in $usersReadyForCompletion) {
                $duration = if ($readyUser.SyncDuration) {
                    "{0:N0}d {1:N0}h {2:N0}m" -f $readyUser.SyncDuration.Days, $readyUser.SyncDuration.Hours, $readyUser.SyncDuration.Minutes
                }
                else {
                    "N/A"
                }

                Write-Host "      • $($readyUser.Identity) - $($readyUser.PercentageComplete)% (trajanje: $duration)" -ForegroundColor White
            }

            Write-Host "`n   ⏳ Preostali mailboxevi još nisu dostigli 95%" -ForegroundColor Yellow
            Write-Host "   💡 Pričekajte da svi mailboxevi dostignu 95% prije finalizacije batcha" -ForegroundColor Cyan
            Write-Host ""
        }
        else {
            Write-Host "⏳ Batch u toku: " -NoNewline -ForegroundColor Cyan
            Write-Host $batch.Identity -ForegroundColor White
            Write-Host "   📊 Nijedan mailbox još nije dostigao 95%" -ForegroundColor Cyan
            Write-Host ""
        }
    }

    # Sažetak
    Write-Host "╔════════════════════════════════════════════════════════════╗" -ForegroundColor Green
    Write-Host "║  Sažetak                                                   ║" -ForegroundColor Green
    Write-Host "╚════════════════════════════════════════════════════════════╝" -ForegroundColor Green

    Write-Host ""
    Write-Host "  Batch-eva provjereno: $($batches.Count)"
    Write-Host "  Batch-eva spremno za finalizaciju: " -NoNewline
    Write-Host $readyBatches.Count -ForegroundColor $(if ($readyBatches.Count -gt 0) { 'Green' } else { 'Yellow' })

    Write-Host "  Mailboxeva spremno za finalizaciju: " -NoNewline
    Write-Host $readyMailboxes.Count -ForegroundColor $(if ($readyMailboxes.Count -gt 0) { 'Green' } else { 'Yellow' })
    Write-Host ""

    # Auto-complete ako je traženo
    if ($AutoComplete -and $readyBatches.Count -gt 0) {
        Write-Host "⚠️  UPOZORENJE: AutoComplete je omogućen!" -ForegroundColor Red
        Write-Host "   Pokrećem finalizaciju za $($readyBatches.Count) batch(eva)...`n" -ForegroundColor Red

        foreach ($batchToComplete in $readyBatches) {
            Write-Host "   ▶️  Finaliziram: $batchToComplete..." -ForegroundColor Yellow

            try {
                Complete-MigrationBatch -Identity $batchToComplete -Confirm:$false -ErrorAction Stop
                Write-Host "   ✅ Uspješno pokrenuta finalizacija za: $batchToComplete" -ForegroundColor Green
            }
            catch {
                Write-Host "   ❌ Greška pri finalizaciji $batchToComplete : $_" -ForegroundColor Red
            }

            Write-Host ""
        }
    }
    elseif ($readyBatches.Count -gt 0) {
        Write-Host "💡 Savjet: Dodajte '-AutoComplete' parametar za automatsku finalizaciju" -ForegroundColor Cyan
        Write-Host "   (OPREZ: Finalizacija će se pokrenuti bez dodatne potvrde!)" -ForegroundColor Yellow
        Write-Host ""
    }

    if ($readyBatches.Count -eq 0) {
        Write-Host "💤 Nijedan batch još nije spreman za finalizaciju." -ForegroundColor Yellow
        Write-Host "   Pričekajte da mailboxevi dostignu 95% completion." -ForegroundColor Cyan
        Write-Host ""
    }
}
catch {
    Write-Error "Greška: $_"
    Write-Error $_.Exception.Message
}
