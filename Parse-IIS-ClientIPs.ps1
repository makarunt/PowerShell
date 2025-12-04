# IIS Log Parser za IMAP/POP3 klijentske IP adrese
# Ova skripta parsira IIS logove na CAS (Client Access Server) gdje se nalazi PRAVA klijentska IP

# 1. Postavke
$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-2)

# 2. Putanja do IIS logova
# Obično su na CAS serveru, ne na Mailbox serveru
$IISLogPath = "C:\inetpub\logs\LogFiles\W3SVC1"  # Može biti W3SVC1, W3SVC2, itd.

# Alternativne putanje (pokušaj sve)
$AlternativePaths = @(
    "C:\inetpub\logs\LogFiles\W3SVC1",
    "C:\inetpub\logs\LogFiles\W3SVC2",
    "C:\inetpub\logs\LogFiles\W3SVC3"
)

# Export putanja
$ExportBase = "C:\Temp\Span\"
if (-not (Test-Path -Path $ExportBase)) {
    New-Item -Path $ExportBase -ItemType Directory | Out-Null
}

Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  IIS Log Parser - Prave Klijentske IP Adrese" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

$AllResults = @()

# Funkcija za parsiranje IIS logova
function Parse-IISLogs {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,

        [Parameter(Mandatory=$true)]
        [ref]$Results
    )

    if (-not (Test-Path -Path $Path)) {
        Write-Warning "Putanja ne postoji: $Path"
        return
    }

    Write-Host "Pretražujem IIS logove u: $Path" -ForegroundColor Yellow

    # Traži .log datoteke u zadanom periodu
    $LogFiles = Get-ChildItem -Path $Path -Filter "*.log" -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $StartDate } |
                Sort-Object LastWriteTime -Descending

    if (-not $LogFiles) {
        Write-Warning "  Nema log datoteka za zadani period."
        return
    }

    Write-Host "  Pronađeno $($LogFiles.Count) log datoteka" -ForegroundColor Green

    $TempUnique = @{}
    $TotalProcessed = 0

    foreach ($File in $LogFiles) {
        Write-Host "  Obrada: $($File.Name)" -ForegroundColor Gray

        try {
            $Content = Get-Content -Path $File.FullName -ErrorAction Stop

            # IIS log format: Traži liniju sa #Fields:
            $HeaderLine = $Content | Where-Object { $_.StartsWith('#Fields:') } | Select-Object -First 1

            if (-not $HeaderLine) {
                Write-Warning "    Nema zaglavlja u $($File.Name)"
                continue
            }

            # Izvuci nazive stupaca
            $ColumnNames = ($HeaderLine -replace '#Fields:\s+', '') -split '\s+'

            # Pronađi indexe ključnih stupaca
            $cIpIndex = [array]::IndexOf($ColumnNames, 'c-ip')
            $csUsernameIndex = [array]::IndexOf($ColumnNames, 'cs-username')
            $csUriStemIndex = [array]::IndexOf($ColumnNames, 'cs-uri-stem')
            $dateIndex = [array]::IndexOf($ColumnNames, 'date')
            $timeIndex = [array]::IndexOf($ColumnNames, 'time')
            $scStatusIndex = [array]::IndexOf($ColumnNames, 'sc-status')

            # Filtriraj podatkovne linije
            $DataLines = $Content | Where-Object {
                -not $_.StartsWith('#') -and
                $_.Trim() -ne '' -and
                # Filtriraj samo IMAP/POP3 relevantne URI-je
                ($_ -match '/Microsoft-Server-ActiveSync' -or
                 $_ -match '/mapi' -or
                 $_ -match '/rpc' -or
                 $_ -match '/ews' -or
                 $_ -match 'authenticated')  # Ili bilo što što pokazuje autentikaciju
            }

            foreach ($Line in $DataLines) {
                $Fields = $Line -split '\s+'

                if ($Fields.Count -lt $ColumnNames.Count) { continue }

                # Izvuci podatke
                $ClientIP = if ($cIpIndex -ge 0) { $Fields[$cIpIndex] } else { '' }
                $Username = if ($csUsernameIndex -ge 0) { $Fields[$csUsernameIndex] } else { '' }
                $UriStem = if ($csUriStemIndex -ge 0) { $Fields[$csUriStemIndex] } else { '' }
                $Date = if ($dateIndex -ge 0) { $Fields[$dateIndex] } else { '' }
                $Time = if ($timeIndex -ge 0) { $Fields[$timeIndex] } else { '' }
                $Status = if ($scStatusIndex -ge 0) { $Fields[$scStatusIndex] } else { '' }

                # Preskoci prazne ili '-' vrijednosti
                if ([string]::IsNullOrWhiteSpace($Username) -or $Username -eq '-') { continue }
                if ([string]::IsNullOrWhiteSpace($ClientIP) -or $ClientIP -eq '-') { continue }

                # Filtriranje HealthMailbox
                if ($Username -match 'HealthMailbox') { continue }

                # Provjera datuma
                if ($Date -and $Date -ne '-') {
                    try {
                        $LogDate = [DateTime]::ParseExact("$Date $Time", 'yyyy-MM-dd HH:mm:ss', $null)
                        if ($LogDate -lt $StartDate) { continue }
                    }
                    catch {
                        continue
                    }
                }

                $TotalProcessed++

                # Deduplikacija
                $Key = "$Username|$ClientIP"
                if (-not $TempUnique.ContainsKey($Key)) {
                    $TempUnique.Add($Key,
                        [PSCustomObject]@{
                            UserName = $Username
                            ClientIP = $ClientIP
                            LastSeen = "$Date $Time"
                            LastURI = $UriStem
                            LastStatus = $Status
                        }
                    )
                }
            }

            Write-Host "    Procesuirano: $TotalProcessed zapisa" -ForegroundColor DarkGray

        }
        catch {
            Write-Warning "    Greška: $($_.Exception.Message)"
        }
    }

    Write-Host "  Pronađeno $($TempUnique.Count) unikatnih User/IP parova" -ForegroundColor Green

    # Dodaj rezultate
    $Results.Value += $TempUnique.Values
}

# Pokušaj sve moguće putanje
foreach ($LogPath in $AlternativePaths) {
    Parse-IISLogs -Path $LogPath -Results ([ref]$AllResults)
}

# Eksport rezultata
if ($AllResults.Count -gt 0) {
    $ExportPath = "$($ExportBase)IIS_Client_IPs_$($EndDate.ToString('yyyyMMdd_HHmmss')).csv"
    $AllResults |
        Select-Object UserName, ClientIP, LastSeen, LastURI, LastStatus |
        Sort-Object UserName, ClientIP |
        Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

    Write-Host "`n✅ IIS Log izvještaj spremljen: $ExportPath" -ForegroundColor Green
    Write-Host "   Pronađeno $($AllResults.Count) unikatnih autentikacija" -ForegroundColor Green

    # Prikaži prvih 10 rezultata
    Write-Host "`n─────────────────────────────────────────────────────────" -ForegroundColor Yellow
    Write-Host "  Primjer rezultata (prvih 10):" -ForegroundColor Yellow
    Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor Yellow
    $AllResults | Select-Object -First 10 | Format-Table -AutoSize

} else {
    Write-Host "`n⚠️  Nema pronađenih autentikacija u IIS logovima." -ForegroundColor Yellow
    Write-Host @"

NAPOMENA: Ova skripta traži IIS logove na OVOM serveru.
Ako je ovo Mailbox server, IIS logovi sa pravim klijentskim IP adresama
nalaze se na CAS (Client Access Server) serveru.

Potrebno je:
1. Pokrenuti ovu skriptu na CAS serveru, ILI
2. Pristupiti IIS logovima na CAS serveru putem network share-a

Tipična putanja na CAS serveru:
  \\CAS-SERVER\c$\inetpub\logs\LogFiles\W3SVC1\

"@ -ForegroundColor Yellow
}

Write-Host "`n═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
