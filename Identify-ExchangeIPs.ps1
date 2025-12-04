# Exchange IP Identification Tool
# Identific pronađene IP adrese iz logova - da li su to serveri ili klijenti?

param(
    [Parameter(Mandatory=$false)]
    [string]$LogPath = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4",

    [Parameter(Mandatory=$false)]
    [int]$DaysBack = 2
)

$EndDate = Get-Date
$StartDate = $EndDate.AddDays(-$DaysBack)

Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Exchange IP Identification Tool" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host ""

# 1. Prvo prikupimo sve IP adrese iz logova
Write-Host "Korak 1: Prikupljam IP adrese iz logova..." -ForegroundColor Yellow

$UniqueIPs = @{}

$LogFiles = Get-ChildItem -Path $LogPath -Filter "*.LOG" -ErrorAction SilentlyContinue |
            Where-Object { $_.CreationTime -ge $StartDate } |
            Select-Object -First 5

if (-not $LogFiles) {
    Write-Error "Nema log datoteka u $LogPath"
    exit
}

foreach ($File in $LogFiles) {
    $Content = Get-Content -Path $File.FullName -ErrorAction SilentlyContinue

    $HeaderLine = $Content | Where-Object { $_.StartsWith('#Fields:') } | Select-Object -First 1
    if (-not $HeaderLine) { continue }

    $ColumnNames = ($HeaderLine -replace '#Fields:', '').Trim()

    $DataContent = $Content | Where-Object {
        -not $_.StartsWith('#') -and
        -not $_.Contains("HealthMailbox") -and
        $_.Contains(',')
    }

    if (-not $DataContent) { continue }

    $CSVInput = @($ColumnNames) + $DataContent
    $ImportedData = $CSVInput | ConvertFrom-Csv -Delimiter ',' -ErrorAction SilentlyContinue

    # Prikupi sve IP adrese iz cIp i sIp polja
    foreach ($Entry in $ImportedData) {
        if ($Entry.cIp) {
            $IP = ($Entry.cIp -replace ':\d+$', '' -replace '\]:\d+$', ']')
            if ($UniqueIPs.ContainsKey($IP)) {
                $UniqueIPs[$IP]++
            } else {
                $UniqueIPs.Add($IP, 1)
            }
        }

        if ($Entry.sIp) {
            $IP = ($Entry.sIp -replace ':\d+$', '' -replace '\]:\d+$', ']')
            if ($UniqueIPs.ContainsKey($IP)) {
                $UniqueIPs[$IP]++
            } else {
                $UniqueIPs.Add($IP, 1)
            }
        }
    }
}

Write-Host "Pronađeno $($UniqueIPs.Count) unikatnih IP adresa u logovima" -ForegroundColor Green
Write-Host ""

# 2. Identifikacija IP adresa
Write-Host "Korak 2: Identifikacija IP adresa..." -ForegroundColor Yellow
Write-Host ""

$Results = @()

foreach ($IP in $UniqueIPs.Keys | Sort-Object) {
    $Count = $UniqueIPs[$IP]
    $CleanIP = $IP -replace '[\[\]]', ''  # Ukloni [ i ] za IPv6

    Write-Host "─────────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host "IP Adresa: $IP (Pojavljivanja: $Count)" -ForegroundColor White

    $Result = [PSCustomObject]@{
        IPAddress = $IP
        Occurrences = $Count
        IsLocalServer = $false
        IsExchangeServer = $false
        Hostname = ""
        IPType = ""
        Description = ""
    }

    # Provjeri da li je lokalni server
    $LocalIPs = Get-NetIPAddress -ErrorAction SilentlyContinue | Select-Object -ExpandProperty IPAddress
    if ($CleanIP -in $LocalIPs) {
        $Result.IsLocalServer = $true
        $Result.Description = "OVAJ SERVER (lokalna IP adresa)"
        Write-Host "  ✓ OVO JE OVAJ SERVER (lokalna IP adresa)" -ForegroundColor Green
    }

    # Provjeri da li je RFC1918 privatna IP (10.x, 172.16-31.x, 192.168.x)
    if ($CleanIP -match '^10\.' -or
        $CleanIP -match '^172\.(1[6-9]|2[0-9]|3[0-1])\.' -or
        $CleanIP -match '^192\.168\.') {
        $Result.IPType = "Privatna (RFC1918)"
        Write-Host "  → Privatna IP adresa (RFC1918) - vjerovatno interni server" -ForegroundColor Yellow
    }
    elseif ($CleanIP -match '^169\.254\.') {
        $Result.IPType = "APIPA"
        Write-Host "  → APIPA adresa (169.254.x.x)" -ForegroundColor DarkYellow
    }
    else {
        $Result.IPType = "Javna"
        Write-Host "  → Javna IP adresa - vjerovatno eksterni klijent" -ForegroundColor Cyan
    }

    # Pokušaj DNS rezoluciju
    try {
        $DNS = Resolve-DnsName -Name $CleanIP -ErrorAction Stop | Select-Object -First 1
        if ($DNS.NameHost) {
            $Result.Hostname = $DNS.NameHost
            Write-Host "  → Hostname: $($DNS.NameHost)" -ForegroundColor White

            if ($DNS.NameHost -match 'exchange|cas|mail|mex|hub') {
                $Result.IsExchangeServer = $true
                $Result.Description += " Exchange server"
                Write-Host "    ⚠️ IZGLEDA KAO EXCHANGE SERVER!" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "  → Nema DNS zapisa" -ForegroundColor DarkGray
    }

    # Pokušaj ping (provjeri da li je dostupan)
    $Ping = Test-Connection -ComputerName $CleanIP -Count 1 -Quiet -ErrorAction SilentlyContinue
    if ($Ping) {
        Write-Host "  → Server je ONLINE" -ForegroundColor Green
    } else {
        Write-Host "  → Server je OFFLINE ili blokira ping" -ForegroundColor DarkGray
    }

    # Pokušaj dohvatiti Exchange server info (ako je dostupan)
    if ($Result.IsLocalServer) {
        try {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue
            $ExchangeServers = Get-ExchangeServer -ErrorAction SilentlyContinue

            foreach ($Server in $ExchangeServers) {
                $ServerIPs = [System.Net.Dns]::GetHostAddresses($Server.Fqdn) | Select-Object -ExpandProperty IPAddressToString
                if ($CleanIP -in $ServerIPs) {
                    $Result.IsExchangeServer = $true
                    $Result.Description = "Exchange Server: $($Server.Name) ($($Server.ServerRole))"
                    Write-Host "  ✓ Exchange Server: $($Server.Name)" -ForegroundColor Magenta
                    Write-Host "    Role: $($Server.ServerRole)" -ForegroundColor Magenta
                }
            }
        }
        catch {
            # Nema Exchange cmdlet-a
        }
    }

    $Results += $Result
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  SAŽETAK" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan

# Grupiraj rezultate
$LocalServers = $Results | Where-Object { $_.IsLocalServer }
$ExchangeServers = $Results | Where-Object { $_.IsExchangeServer }
$PrivateIPs = $Results | Where-Object { $_.IPType -eq "Privatna (RFC1918)" -and -not $_.IsLocalServer }
$PublicIPs = $Results | Where-Object { $_.IPType -eq "Javna" }

Write-Host ""
Write-Host "Lokalni serveri (ovaj server):" -ForegroundColor Green
if ($LocalServers) {
    $LocalServers | Format-Table IPAddress, Occurrences, Description -AutoSize
} else {
    Write-Host "  Nema" -ForegroundColor DarkGray
}

Write-Host "Exchange serveri:" -ForegroundColor Magenta
if ($ExchangeServers) {
    $ExchangeServers | Format-Table IPAddress, Occurrences, Hostname, Description -AutoSize
} else {
    Write-Host "  Nema identifikovanih" -ForegroundColor DarkGray
}

Write-Host "Privatne IP adrese (potencijalni interni serveri):" -ForegroundColor Yellow
if ($PrivateIPs) {
    $PrivateIPs | Format-Table IPAddress, Occurrences, Hostname -AutoSize
    Write-Host "  ⚠️ Ovo su vjerovatno CAS serveri ili load balanceri!" -ForegroundColor Red
} else {
    Write-Host "  Nema" -ForegroundColor DarkGray
}

Write-Host "Javne IP adrese (potencijalni eksterni klijenti):" -ForegroundColor Cyan
if ($PublicIPs) {
    $PublicIPs | Format-Table IPAddress, Occurrences, Hostname -AutoSize
    Write-Host "  ✓ Ovo su vjerovatno PRAVE klijentske IP adrese!" -ForegroundColor Green
} else {
    Write-Host "  Nema" -ForegroundColor DarkGray
}

Write-Host ""
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  ZAKLJUČAK" -ForegroundColor Cyan
Write-Host "═══════════════════════════════════════════════════════════" -ForegroundColor Cyan

if ($PrivateIPs.Count -gt 0) {
    Write-Host @"

⚠️  PRONAĐENE SU PRIVATNE IP ADRESE U 'cIp' POLJU!

Ovo znači da logovi koje gledate (na Mailbox serveru) NE SADRŽE
prave klijentske IP adrese, već IP adrese CAS servera ili load balancera.

RJEŠENJE:
1. Provjerite IMAP/POP3 logove NA CAS SERVERU (Frontend)
   Putanja: \\CAS-SERVER\C$\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4
   Putanja: \\CAS-SERVER\C$\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3

2. CAS server prima direktne konekcije od klijenata i tamo bi 'cIp'
   polje trebalo sadržavati PRAVE klijentske IP adrese

3. Pokrenite Parse-ExchangeMailLogs-Fixed.ps1 NA CAS SERVERU

"@ -ForegroundColor Yellow
} elseif ($PublicIPs.Count -gt 0) {
    Write-Host @"

✅ PRONAĐENE SU JAVNE IP ADRESE!

Ovo su vjerovatno PRAVE klijentske IP adrese.
Možete sigurno koristiti Parse-ExchangeMailLogs-Fixed.ps1 skriptu.

"@ -ForegroundColor Green
} else {
    Write-Host @"

ℹ️  Svi logovi sadrže samo lokalne server IP adrese.

Provjerite da li gledate prave logove i da li ima autentikacija
u zadnjih $DaysBack dana.

"@ -ForegroundColor White
}

# Eksport rezultata
$ExportPath = "C:\Temp\Span\IP_Identification_Report.csv"
$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -ErrorAction SilentlyContinue
if (Test-Path $ExportPath) {
    Write-Host "Detaljan izvještaj spremljen: $ExportPath" -ForegroundColor Cyan
}
