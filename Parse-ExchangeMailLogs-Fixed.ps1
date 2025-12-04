# Exchange IMAP/POP3 Log Parser - Fixed Version
# Uzima IP adrese SAMO iz autentikacijskih komandi

# 1. Postavite datumski raspon (zadnjih x dana)
$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-2)

# 2. FIKSNA PUTANJA ZA EKSPORT
$ExportBase = "C:\Temp\Span\"

# 3. Putanje do logova
$ImapLogPath = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4"
$Pop3LogPath = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3"

# Provjera postojanja mape za eksport
if (-not (Test-Path -Path $ExportBase)) {
    New-Item -Path $ExportBase -ItemType Directory | Out-Null
    Write-Host "Kreirana mapa za eksport: $ExportBase" -ForegroundColor Yellow
}

$ImapUniqueResults = @()
$Pop3UniqueResults = @()

# Funkcija za obradu logova - uzima samo autentikacijske komande
function Get-ProtocolLogData-AuthOnly {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,

        [Parameter(Mandatory=$true)]
        [string]$Protocol,

        [Parameter(Mandatory=$true)]
        [string]$AuthCommand, # "authenticate" za IMAP, "auth" za POP3

        [Parameter(Mandatory=$true)]
        [ref]$ResultList
    )

    Write-Host "Pretražujem logove za protokol $Protocol (samo autentikacijske komande)..." -ForegroundColor Yellow

    $LogFiles = Get-ChildItem -Path $Path -Filter "*.LOG" -ErrorAction SilentlyContinue |
                Where-Object { $_.CreationTime -ge $StartDate }

    if (-not $LogFiles) {
        Write-Warning "Nema log datoteka u $Path za zadani period."
        return
    }

    $TempUnique = @{}
    $ProcessedCount = 0
    $AuthCount = 0

    foreach ($File in $LogFiles) {
        Write-Host "  Obrada: $($File.Name)" -ForegroundColor Gray

        try {
            # Učitaj cijeli sadržaj datoteke
            $Content = Get-Content -Path $File.FullName -ErrorAction Stop

            # Pronađi liniju zaglavlja (#Fields:)
            $HeaderLine = $Content | Where-Object { $_.StartsWith('#Fields:') } | Select-Object -First 1

            if (-not $HeaderLine) {
                Write-Warning "    Preskačem $($File.Name): Nema zaglavlja (#Fields:)"
                continue
            }

            # Izvuci nazive stupaca
            $ColumnNames = ($HeaderLine -replace '#Fields:', '').Trim()

            # Filtriraj samo podatkovne linije (bez komentara i HealthMailbox)
            $DataContent = $Content |
                Where-Object {
                    -not $_.StartsWith('#') -and
                    -not $_.Contains("HealthMailbox") -and
                    $_.Contains(',')
                }

            if (-not $DataContent) {
                Write-Host "    Nema podataka za obradu." -ForegroundColor Gray
                continue
            }

            # Konvertiraj u CSV objekte
            $CSVInput = @($ColumnNames) + $DataContent
            $ImportedData = $CSVInput | ConvertFrom-Csv -Delimiter ',' -ErrorAction Stop

            # Filtriraj samo autentikacijske komande
            $AuthEntries = $ImportedData | Where-Object {
                $_.command -eq $AuthCommand -and
                -not [string]::IsNullOrWhiteSpace($_.user) -and
                -not [string]::IsNullOrWhiteSpace($_.cIp)
            }

            $ProcessedCount += $AuthEntries.Count

            foreach ($Entry in $AuthEntries) {

                # Provjeri datum
                try {
                    $LogDate = [DateTime]::ParseExact($Entry.dateTime.Substring(0, 19), 'yyyy-MM-ddTHH:mm:ss', $null)
                }
                catch {
                    continue
                }

                if ($LogDate -lt $StartDate) { continue }

                $User = $Entry.user.Trim()
                $IP = $Entry.cIp.Trim()

                # Ukloni port iz IP adrese
                if ($IP -match "^\[.*\]:\d+$") {
                    # IPv6 format [xxxx:xxxx]:port
                    $IP = $IP -replace ":\d+$", ""
                } elseif ($IP -match "^[\d\.]+:\d+$") {
                    # IPv4 format xxx.xxx.xxx.xxx:port
                    $IP = $IP -replace ":\d+$", ""
                }

                # Deduplikacija: Username + IP
                $Key = "$User|$IP"
                if (-not $TempUnique.ContainsKey($Key)) {
                    $TempUnique.Add($Key,
                        [PSCustomObject]@{
                            UserName = $User
                            IPAddress = $IP
                            Protocol = $Protocol
                            LastSeen = $LogDate
                        }
                    )
                    $AuthCount++
                }
            }
        }
        catch {
            Write-Warning "    Greška pri obradi $($File.Name): $($_.Exception.Message)"
        }
    }

    Write-Host "  Pronađeno: $ProcessedCount autentikacijskih zapisa, $AuthCount unikatnih parova User/IP" -ForegroundColor Green

    # Dodaj sve unikatne rezultate u konačnu listu
    $ResultList.Value += $TempUnique.Values
}

# 4. Izvršavanje funkcije za oba protokola

Write-Host "`n=== IMAP Analiza ===" -ForegroundColor Cyan
Get-ProtocolLogData-AuthOnly -Path $ImapLogPath -Protocol "IMAP" -AuthCommand "authenticate" -ResultList ([ref]$ImapUniqueResults)

Write-Host "`n=== POP3 Analiza ===" -ForegroundColor Cyan
Get-ProtocolLogData-AuthOnly -Path $Pop3LogPath -Protocol "POP3" -AuthCommand "auth" -ResultList ([ref]$Pop3UniqueResults)

# 5. Eksportiranje rezultata

Write-Host "`n=== Spremanje rezultata ===" -ForegroundColor Yellow

# --- IMAP Izvještaj ---
if ($ImapUniqueResults.Count -gt 0) {
    $ImapExportPath = "$($ExportBase)Exchange_IMAP_Unique_Auth_$($EndDate.ToString('yyyyMMdd')).csv"
    $ImapUniqueResults |
        Select-Object UserName, IPAddress, @{Name="LastSeen";Expression={$_.LastSeen.ToString('yyyy-MM-dd HH:mm:ss')}} |
        Sort-Object UserName, IPAddress |
        Export-Csv -Path $ImapExportPath -NoTypeInformation -Encoding UTF8

    Write-Host "✅ IMAP Izvještaj: $ImapExportPath" -ForegroundColor Green
    Write-Host "   Pronađeno $($ImapUniqueResults.Count) unikatnih IMAP autentikacija" -ForegroundColor Green
} else {
    Write-Host "⚠️  Nema IMAP autentikacija za zadani period." -ForegroundColor Yellow
}

# --- POP3 Izvještaj ---
if ($Pop3UniqueResults.Count -gt 0) {
    $Pop3ExportPath = "$($ExportBase)Exchange_POP3_Unique_Auth_$($EndDate.ToString('yyyyMMdd')).csv"
    $Pop3UniqueResults |
        Select-Object UserName, IPAddress, @{Name="LastSeen";Expression={$_.LastSeen.ToString('yyyy-MM-dd HH:mm:ss')}} |
        Sort-Object UserName, IPAddress |
        Export-Csv -Path $Pop3ExportPath -NoTypeInformation -Encoding UTF8

    Write-Host "✅ POP3 Izvještaj: $Pop3ExportPath" -ForegroundColor Green
    Write-Host "   Pronađeno $($Pop3UniqueResults.Count) unikatnih POP3 autentikacija" -ForegroundColor Green
} else {
    Write-Host "⚠️  Nema POP3 autentikacija za zadani period." -ForegroundColor Yellow
}

# --- Kombinirani izvještaj (opciono) ---
$AllResults = $ImapUniqueResults + $Pop3UniqueResults
if ($AllResults.Count -gt 0) {
    $CombinedExportPath = "$($ExportBase)Exchange_Combined_Unique_Auth_$($EndDate.ToString('yyyyMMdd')).csv"
    $AllResults |
        Select-Object UserName, IPAddress, Protocol, @{Name="LastSeen";Expression={$_.LastSeen.ToString('yyyy-MM-dd HH:mm:ss')}} |
        Sort-Object UserName, IPAddress, Protocol |
        Export-Csv -Path $CombinedExportPath -NoTypeInformation -Encoding UTF8

    Write-Host "✅ Kombinirani izvještaj: $CombinedExportPath" -ForegroundColor Green
    Write-Host "   Ukupno $($AllResults.Count) unikatnih autentikacija" -ForegroundColor Green
}

Write-Host "`n=== Završeno ===" -ForegroundColor Cyan
