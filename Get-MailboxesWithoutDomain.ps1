<#
.SYNOPSIS
    Pronalazi mailboxe koji nemaju email adresu na specifičnoj domeni.

.DESCRIPTION
    Ova skripta prolazi kroz sve mailboxe i ispisuje one koji nemaju
    niti jednu email adresu (proxy address) na definisanoj domeni.

    Korisno za identifikaciju mailboxova koji nisu konfigurisani sa
    adresama na određenoj domeni nakon migracije ili reorganizacije.

.PARAMETER Domain
    Domena za pretragu (npr. "xyz.com"). Obavezno.

.PARAMETER ExportToCsv
    Putanja do CSV fajla za export rezultata (opciono).

.PARAMETER IncludeShared
    Uključi shared mailboxe u pretragu.

.PARAMETER IncludeRoom
    Uključi room mailboxe u pretragu.

.PARAMETER IncludeEquipment
    Uključi equipment mailboxe u pretragu.

.EXAMPLE
    .\Get-MailboxesWithoutDomain.ps1 -Domain "xyz.com"

    Prikazuje sve user mailboxe koji nemaju adresu na xyz.com domeni.

.EXAMPLE
    .\Get-MailboxesWithoutDomain.ps1 -Domain "xyz.com" -ExportToCsv "C:\Results\mailboxes.csv"

    Exportuje rezultate u CSV fajl.

.EXAMPLE
    .\Get-MailboxesWithoutDomain.ps1 -Domain "xyz.com" -IncludeShared -IncludeRoom

    Uključuje shared i room mailboxe u pretragu.

.NOTES
    Zahteva:
    - Exchange Management Shell (On-Premises) ili
    - Exchange Online PowerShell Module (Microsoft 365)
    - Odgovarajuće Exchange admin dozvole

    Autor: PowerShell Script
    Verzija: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Unesite domenu za pretragu (npr. xyz.com)")]
    [ValidateNotNullOrEmpty()]
    [string]$Domain,

    [Parameter(Mandatory = $false)]
    [string]$ExportToCsv,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeShared,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeRoom,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeEquipment
)

# Funkcija za provjeru konekcije na Exchange
function Test-ExchangeConnection {
    try {
        $null = Get-Command Get-Mailbox -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Glavna skripta
try {
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Mailbox Domain Filter Analyzer" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan

    # Provjeri Exchange konekciju
    Write-Host "Provjeravam konekciju na Exchange..." -ForegroundColor Yellow
    if (-not (Test-ExchangeConnection)) {
        throw "Nije uspostavljena konekcija na Exchange. Molimo pokrenite Exchange Management Shell ili se konektujte sa Connect-ExchangeOnline."
    }
    Write-Host "✓ Konekcija uspješna`n" -ForegroundColor Green

    # Normalizuj domenu (ukloni @ ako postoji)
    $targetDomain = $Domain.TrimStart('@').ToLower()
    Write-Host "Tražim mailboxe BEZ adresa na domeni: @$targetDomain`n" -ForegroundColor Cyan

    # Pripremi filter za tipove mailboxova
    $recipientTypeDetails = @('UserMailbox')

    if ($IncludeShared) {
        $recipientTypeDetails += 'SharedMailbox'
        Write-Host "  • Uključujem Shared mailboxe" -ForegroundColor Gray
    }
    if ($IncludeRoom) {
        $recipientTypeDetails += 'RoomMailbox'
        Write-Host "  • Uključujem Room mailboxe" -ForegroundColor Gray
    }
    if ($IncludeEquipment) {
        $recipientTypeDetails += 'EquipmentMailbox'
        Write-Host "  • Uključujem Equipment mailboxe" -ForegroundColor Gray
    }

    Write-Host "`nUčitavam mailboxe..." -ForegroundColor Yellow

    # Učitaj sve mailboxe
    $allMailboxes = @()
    foreach ($type in $recipientTypeDetails) {
        Write-Host "  Učitavam $type..." -ForegroundColor Gray
        $mailboxes = Get-Mailbox -RecipientTypeDetails $type -ResultSize Unlimited -ErrorAction Stop
        $allMailboxes += $mailboxes
    }

    $totalCount = $allMailboxes.Count
    Write-Host "✓ Učitano $totalCount mailboxova`n" -ForegroundColor Green

    # Filtriraj mailboxe koji NEMAJU adresu na target domeni
    Write-Host "Analiziram email adrese..." -ForegroundColor Yellow

    $mailboxesWithoutDomain = @()
    $counter = 0

    foreach ($mailbox in $allMailboxes) {
        $counter++

        # Progress bar
        if ($counter % 50 -eq 0 -or $counter -eq $totalCount) {
            $percentComplete = [math]::Round(($counter / $totalCount) * 100)
            Write-Progress -Activity "Analiziram mailboxe" -Status "$counter od $totalCount" -PercentComplete $percentComplete
        }

        # Provjeri da li mailbox ima adresu na target domeni
        $hasTargetDomain = $false

        foreach ($address in $mailbox.EmailAddresses) {
            # Provjeri SMTP adrese (smtp: i SMTP:)
            if ($address -is [string] -and $address -match '^(smtp|SMTP):(.+)@(.+)$') {
                $emailDomain = $Matches[3].ToLower()

                if ($emailDomain -eq $targetDomain) {
                    $hasTargetDomain = $true
                    break
                }
            }
        }

        # Ako mailbox NEMA adresu na target domeni, dodaj ga u rezultate
        if (-not $hasTargetDomain) {
            $mailboxInfo = [PSCustomObject]@{
                DisplayName        = $mailbox.DisplayName
                Alias             = $mailbox.Alias
                PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                UserPrincipalName = $mailbox.UserPrincipalName
                RecipientType     = $mailbox.RecipientTypeDetails
                EmailAddresses    = ($mailbox.EmailAddresses | Where-Object { $_ -like "smtp:*" -or $_ -like "SMTP:*" }) -join "; "
            }

            $mailboxesWithoutDomain += $mailboxInfo
        }
    }

    Write-Progress -Activity "Analiziram mailboxe" -Completed

    # Prikaz rezultata
    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "REZULTATI" -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan

    $foundCount = $mailboxesWithoutDomain.Count
    Write-Host "Pronađeno: $foundCount mailboxa BEZ adrese na @$targetDomain" -ForegroundColor $(if ($foundCount -gt 0) { "Yellow" } else { "Green" })
    Write-Host "Ukupno analizirano: $totalCount mailboxova`n" -ForegroundColor Gray

    if ($foundCount -gt 0) {
        # Prikaži mailboxe
        Write-Host "Mailboxes bez @$targetDomain adrese:`n" -ForegroundColor Yellow

        $mailboxesWithoutDomain | Format-Table -Property DisplayName, PrimarySmtpAddress, RecipientType -AutoSize

        # Export u CSV ako je specificirano
        if ($ExportToCsv) {
            Write-Host "`nExportujem rezultate..." -ForegroundColor Yellow
            $mailboxesWithoutDomain | Export-Csv -Path $ExportToCsv -NoTypeInformation -Encoding UTF8
            Write-Host "✓ Export završen: $ExportToCsv" -ForegroundColor Green
        }

        # Detaljan prikaz (opciono)
        Write-Host "`nŽelite li vidjeti detaljne informacije za sve mailboxe? (y/n): " -ForegroundColor Cyan -NoNewline
        $response = Read-Host

        if ($response -eq 'y' -or $response -eq 'Y') {
            Write-Host "`nDetaljne informacije:`n" -ForegroundColor Cyan
            $mailboxesWithoutDomain | Format-List -Property DisplayName, Alias, PrimarySmtpAddress, UserPrincipalName, RecipientType, EmailAddresses
        }
    }
    else {
        Write-Host "✓ Svi mailboxes imaju bar jednu adresu na @$targetDomain domeni!" -ForegroundColor Green
    }

    Write-Host "`n========================================`n" -ForegroundColor Cyan
}
catch {
    Write-Host "`n✗ GREŠKA: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "`nStack Trace:" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    exit 1
}
