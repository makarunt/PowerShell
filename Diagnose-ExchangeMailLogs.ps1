# Exchange IMAP/POP3 Log Diagnostic - Prikazuje SVA polja iz autentikacijskih zapisa
# Koristi ovu skriptu da vidiš gdje se nalazi prava klijentska IP adresa

# 1. Postavke
$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-2)
$MaxSamples = 10  # Koliko uzoraka prikazati

# 2. Putanje do logova
$ImapLogPath = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4"
$Pop3LogPath = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3"

# Export putanja
$ExportBase = "C:\Temp\Span\"
if (-not (Test-Path -Path $ExportBase)) {
    New-Item -Path $ExportBase -ItemType Directory | Out-Null
}

Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Exchange Mail Log Diagnostic Tool" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# Funkcija za dijagnostičku analizu
function Show-AuthenticationDetails {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path,

        [Parameter(Mandatory=$true)]
        [string]$Protocol,

        [Parameter(Mandatory=$true)]
        [string]$AuthCommand
    )

    Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor Yellow
    Write-Host "  Analiziram $Protocol logove..." -ForegroundColor Yellow
    Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor Yellow

    $LogFiles = Get-ChildItem -Path $Path -Filter "*.LOG" -ErrorAction SilentlyContinue |
                Where-Object { $_.CreationTime -ge $StartDate } |
                Sort-Object CreationTime -Descending |
                Select-Object -First 5  # Samo zadnjih 5 datoteka

    if (-not $LogFiles) {
        Write-Warning "Nema log datoteka u $Path"
        return
    }

    $SampleCount = 0
    $AllAuthRecords = @()

    foreach ($File in $LogFiles) {
        if ($SampleCount -ge $MaxSamples) { break }

        Write-Host "`nČitam: $($File.Name)" -ForegroundColor Gray

        try {
            $Content = Get-Content -Path $File.FullName -ErrorAction Stop

            # Pronađi zaglavlje
            $HeaderLine = $Content | Where-Object { $_.StartsWith('#Fields:') } | Select-Object -First 1
            if (-not $HeaderLine) { continue }

            $ColumnNames = ($HeaderLine -replace '#Fields:', '').Trim()

            # Filtriraj podatke
            $DataContent = $Content |
                Where-Object {
                    -not $_.StartsWith('#') -and
                    -not $_.Contains("HealthMailbox") -and
                    $_.Contains(',')
                }

            if (-not $DataContent) { continue }

            # Konvertiraj u CSV
            $CSVInput = @($ColumnNames) + $DataContent
            $ImportedData = $CSVInput | ConvertFrom-Csv -Delimiter ',' -ErrorAction Stop

            # Filtriraj samo autentikacijske komande
            $AuthEntries = $ImportedData |
                Where-Object {
                    $_.command -eq $AuthCommand -and
                    -not [string]::IsNullOrWhiteSpace($_.user)
                } |
                Select-Object -First ($MaxSamples - $SampleCount)

            foreach ($Entry in $AuthEntries) {
                $AllAuthRecords += $Entry
                $SampleCount++

                Write-Host "`n┌─ Autentikacijski Zapis #$SampleCount ────────────────────" -ForegroundColor Green
                Write-Host "│"

                # Prikaži SVA polja
                $Entry.PSObject.Properties | ForEach-Object {
                    $Name = $_.Name
                    $Value = $_.Value

                    # Highlightaj ključna polja
                    if ($Name -eq "cIp" -or $Name -eq "sIp") {
                        Write-Host "│  $($Name.PadRight(15)) : $Value" -ForegroundColor Cyan
                    }
                    elseif ($Name -eq "user") {
                        Write-Host "│  $($Name.PadRight(15)) : $Value" -ForegroundColor Yellow
                    }
                    elseif ($Name -eq "context") {
                        # Context može biti vrlo dug, prikaži ga posebno
                        Write-Host "│  $($Name.PadRight(15)) : [Prikazano niže]" -ForegroundColor Magenta
                    }
                    else {
                        Write-Host "│  $($Name.PadRight(15)) : $Value" -ForegroundColor White
                    }
                }

                # Prikaži context polje detaljnije (može sadržavati IP informacije)
                if ($Entry.context) {
                    Write-Host "│"
                    Write-Host "│  CONTEXT ANALIZA:" -ForegroundColor Magenta

                    # Traži IP adrese u context polju
                    $IpMatches = [regex]::Matches($Entry.context, '\b(?:\d{1,3}\.){3}\d{1,3}\b')
                    if ($IpMatches.Count -gt 0) {
                        Write-Host "│    ⚠️ PRONAĐENE IP ADRESE U CONTEXT POLJU:" -ForegroundColor Red
                        foreach ($Match in $IpMatches) {
                            Write-Host "│      → $($Match.Value)" -ForegroundColor Red
                        }
                    }

                    # Prikaži context u čitljivijem formatu
                    $ContextParts = $Entry.context -split ';' | Where-Object { $_ -and $_.Trim() }
                    foreach ($Part in $ContextParts) {
                        if ($Part.Length -gt 100) {
                            Write-Host "│    $($Part.Substring(0, 97))..." -ForegroundColor DarkGray
                        } else {
                            Write-Host "│    $Part" -ForegroundColor DarkGray
                        }
                    }
                }

                Write-Host "└────────────────────────────────────────────────────────" -ForegroundColor Green

                if ($SampleCount -ge $MaxSamples) { break }
            }

        }
        catch {
            Write-Warning "Greška pri obradi $($File.Name): $($_.Exception.Message)"
        }
    }

    # Spremi detalje u CSV za daljnju analizu
    if ($AllAuthRecords.Count -gt 0) {
        $DiagnosticPath = "$ExportBase$($Protocol)_Diagnostic_$($EndDate.ToString('yyyyMMdd_HHmmss')).csv"
        $AllAuthRecords | Export-Csv -Path $DiagnosticPath -NoTypeInformation -Encoding UTF8
        Write-Host "`n✅ Diagnostički CSV spremljen: $DiagnosticPath" -ForegroundColor Green
    }

    Write-Host "`nProcesuirano: $SampleCount autentikacijskih zapisa" -ForegroundColor Green
}

# Pokreni dijagnostiku
Show-AuthenticationDetails -Path $ImapLogPath -Protocol "IMAP" -AuthCommand "authenticate"
Write-Host "`n"
Show-AuthenticationDetails -Path $Pop3LogPath -Protocol "POP3" -AuthCommand "auth"

Write-Host "`n═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  VAŽNO: Mogući izvori klijentske IP adrese" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host @"

Ako je 'cIp' polje Exchange server IP (npr. CAS/Load Balancer),
prava klijentska IP može biti u:

1. IIS LOGOVIMA na CAS serveru
   Lokacija: C:\inetpub\logs\LogFiles\W3SVC*\
   Traži: X-Forwarded-For header ili c-ip polje

2. EXCHANGE CAS (Frontend) LOGOVIMA
   Lokacija: C:\Program Files\Microsoft\Exchange Server\V15\Logging\
   Potraži: HttpProxy logove

3. LOAD BALANCER LOGOVIMA
   Ako koristite F5, HAProxy, ili drugi load balancer

4. FIREWALL/PROXY LOGOVIMA
   Ako postoji firewall između klijenta i Exchange-a

═══════════════════════════════════════════════════════════
"@ -ForegroundColor Yellow

Write-Host "`nZa detaljnu analizu, provjeri generirane CSV datoteke u $ExportBase" -ForegroundColor Cyan
Write-Host ""
