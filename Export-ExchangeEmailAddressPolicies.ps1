<#
.SYNOPSIS
    Dokumentira Email Address Policies na on-prem Exchange serveru i exporta u CSV.

.DESCRIPTION
    Skripta dohvaća sve Email Address Policies, za svaku policy prikazuje:
      - Naziv i prioritet
      - Na koga se primjenjuje (filter uvjeti, tip primatelja)
      - Koje email adrese kreira (predlošci adresa)
    Output se sprema u CSV (delimiter ";") koji se može otvoriti u Excelu ili Wordu.

.PARAMETER OutputPath
    Putanja do izlaznog CSV fajla. Default: .\EmailAddressPolicies_<datum>.csv

.PARAMETER IncludeRecipientCount
    Ako se navede, skripta će prebrojati korisnike na koje se policy primjenjuje.
    Može biti sporije u velikim okruženjima.

.EXAMPLE
    .\Export-ExchangeEmailAddressPolicies.ps1

.EXAMPLE
    .\Export-ExchangeEmailAddressPolicies.ps1 -OutputPath "C:\Reports\EAP.csv" -IncludeRecipientCount

.NOTES
    Pokrenuti iz Exchange Management Shell ili dodati Exchange snap-in/modul.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [string]$OutputPath = ".\EmailAddressPolicies_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter()]
    [switch]$IncludeRecipientCount
)

#region --- Provjera Exchange okoline ---

$exchangeLoaded = $false

# Provjera je li Exchange Management Shell već učitan
if (Get-Command Get-EmailAddressPolicy -ErrorAction SilentlyContinue) {
    $exchangeLoaded = $true
}

# Pokušaj učitavanja Exchange snap-ina (Exchange 2010/2013/2016/2019)
if (-not $exchangeLoaded) {
    if (Get-PSSnapin -Registered | Where-Object { $_.Name -eq 'Microsoft.Exchange.Management.PowerShell.SnapIn' }) {
        try {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
            $exchangeLoaded = $true
            Write-Verbose "Exchange snap-in uspješno učitan."
        }
        catch {
            Write-Warning "Nije moguće učitati Exchange snap-in: $_"
        }
    }
}

# Pokušaj spajanja na Exchange putem implicit remoting (Exchange 2013+)
if (-not $exchangeLoaded) {
    Write-Host "Exchange cmdlets nisu pronađeni lokalno." -ForegroundColor Yellow
    $exchangeServer = Read-Host "Unesi FQDN Exchange servera (npr. mail.domena.local) ili pritisni Enter za prekid"

    if ([string]::IsNullOrWhiteSpace($exchangeServer)) {
        Write-Error "Exchange server nije naveden. Skriptu pokreni iz Exchange Management Shell."
        exit 1
    }

    try {
        $session = New-PSSession -ConfigurationName Microsoft.Exchange `
            -ConnectionUri "http://$exchangeServer/PowerShell/" `
            -Authentication Kerberos `
            -ErrorAction Stop

        Import-PSSession $session -DisableNameChecking -AllowClobber | Out-Null
        $exchangeLoaded = $true
        Write-Verbose "Uspješno spojeno na Exchange server: $exchangeServer"
    }
    catch {
        Write-Error "Nije moguće spojiti se na Exchange server '$exchangeServer': $_"
        exit 1
    }
}

#endregion

#region --- Pomoćne funkcije ---

function Format-TemplateList {
    <#
        Pretvara listu predložaka email adresa u čitljiv string.
        Označava primarnu SMTP adresu s oznakom [PRIMARY].
    #>
    param([object[]]$Templates)

    if (-not $Templates) { return "N/A" }

    $formatted = foreach ($t in $Templates) {
        $str = $t.ToString()
        if ($str -cmatch '^SMTP:') {
            "$str [PRIMARY]"
        }
        else {
            $str
        }
    }
    return $formatted -join " | "
}

function Format-ConditionalAttributes {
    <#
        Gradi čitljiv opis uvjeta primjene iz conditional atributa policy-a.
    #>
    param($Policy)

    $parts = [System.Collections.Generic.List[string]]::new()

    # Tip primatelja (mailboxes, mail users, contacts, groups...)
    if ($Policy.IncludedRecipients -and $Policy.IncludedRecipients -ne 'None') {
        $parts.Add("Tip primatelja: $($Policy.IncludedRecipients)")
    }

    # Standardni conditional atributi
    $conditionals = [ordered]@{
        ConditionalDepartment        = "Odjel"
        ConditionalCompany           = "Tvrtka"
        ConditionalStateOrProvince   = "Država/Pokrajina"
        ConditionalCustomAttribute1  = "CustomAttr1"
        ConditionalCustomAttribute2  = "CustomAttr2"
        ConditionalCustomAttribute3  = "CustomAttr3"
        ConditionalCustomAttribute4  = "CustomAttr4"
        ConditionalCustomAttribute5  = "CustomAttr5"
        ConditionalCustomAttribute6  = "CustomAttr6"
        ConditionalCustomAttribute7  = "CustomAttr7"
        ConditionalCustomAttribute8  = "CustomAttr8"
        ConditionalCustomAttribute9  = "CustomAttr9"
        ConditionalCustomAttribute10 = "CustomAttr10"
        ConditionalCustomAttribute11 = "CustomAttr11"
        ConditionalCustomAttribute12 = "CustomAttr12"
        ConditionalCustomAttribute13 = "CustomAttr13"
        ConditionalCustomAttribute14 = "CustomAttr14"
        ConditionalCustomAttribute15 = "CustomAttr15"
    }

    foreach ($attr in $conditionals.GetEnumerator()) {
        $val = $Policy.($attr.Key)
        if ($val) {
            $parts.Add("$($attr.Value): $($val -join ', ')")
        }
    }

    if ($parts.Count -eq 0) { return $null }
    return $parts -join " | "
}

#endregion

#region --- Dohvaćanje i obrada Email Address Policies ---

Write-Host "`nDohvaćam Email Address Policies..." -ForegroundColor Cyan

try {
    $policies = Get-EmailAddressPolicy -ErrorAction Stop | Sort-Object Priority
}
catch {
    Write-Error "Greška pri dohvaćanju Email Address Policies: $_"
    exit 1
}

if (-not $policies) {
    Write-Warning "Nisu pronađene Email Address Policies."
    exit 0
}

Write-Host "Pronađeno $($policies.Count) polic(y/ies). Obrađujem..." -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($policy in $policies) {

    # --- Tip filtera i uvjeti primjene ---
    $filterType    = $policy.RecipientFilterType   # Precanned ili Custom
    $ldapFilter    = if ($policy.RecipientFilter) { $policy.RecipientFilter.ToString() } else { "N/A" }

    # Čitljivi uvjeti (za Precanned filtere)
    $conditionsFriendly = Format-ConditionalAttributes -Policy $policy

    # Ako nema uvjeta i filter je Precanned, policy se primjenjuje na sve
    if ($filterType -eq 'Precanned' -and -not $conditionsFriendly) {
        $conditionsFriendly = "Svi primatelji (nema dodatnih uvjeta)"
    }
    elseif ($filterType -eq 'Custom') {
        $conditionsFriendly = "Prilagođeni filter (vidi LDAP Filter stupac)"
    }

    # --- Broj primatelja (opcijsko) ---
    $recipientCount = "N/A"
    if ($IncludeRecipientCount) {
        try {
            $count = (Get-Recipient -RecipientPreviewFilter $policy.RecipientFilter -ErrorAction Stop |
                      Measure-Object).Count
            $recipientCount = $count
        }
        catch {
            $recipientCount = "Greška pri dohvaćanju"
        }
    }

    # --- Email adresni predlošci ---
    $templates = Format-TemplateList -Templates $policy.EnabledEmailAddressTemplates

    # --- Gradi red za CSV ---
    $row = [PSCustomObject]@{
        "Prioritet"                  = $policy.Priority
        "Naziv Policy-a"             = $policy.Name
        "Tip Filtera"                = $filterType
        "Uvjeti Primjene"            = $conditionsFriendly
        "LDAP Filter (Custom)"       = if ($filterType -eq 'Custom') { $ldapFilter } else { "" }
        "Predlošci Email Adresa"     = $templates
        "Broj Primatelja"            = $recipientCount
        "Zadnja Primjena"            = if ($policy.LastUpdatedRecipientFilter) {
                                           $policy.LastUpdatedRecipientFilter.ToString("yyyy-MM-dd HH:mm")
                                       } else { "Nikad" }
        "Exchange Organizacija"      = $policy.OrganizationId
    }

    $results.Add($row)

    Write-Verbose "Obrađen: $($policy.Name) [Prioritet: $($policy.Priority)]"
}

#endregion

#region --- Export u CSV ---

try {
    # Osiguraj da direktorij postoji
    $outputDir = Split-Path -Path $OutputPath -Parent
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    Write-Host "`n[OK] Export uspješan: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
    Write-Host "     Broj redova: $($results.Count)" -ForegroundColor Green
    Write-Host "`nUputa za import u Excel:" -ForegroundColor Yellow
    Write-Host "  Podaci > Iz teksta/CSV > odaberi fajl > Delimiter: Točka-zarez (;)" -ForegroundColor Yellow
}
catch {
    Write-Error "Greška pri exportu u CSV: $_"
    exit 1
}

#endregion

# Vrati objekte u pipeline za eventualno daljnje procesiranje
$results
