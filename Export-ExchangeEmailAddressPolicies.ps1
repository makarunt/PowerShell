<#
.SYNOPSIS
    Dokumentira Email Address Policies na on-prem Exchange serveru i exporta u CSV.

.DESCRIPTION
    Skripta dohvaca sve Email Address Policies, za svaku policy prikazuje:
      - Naziv i prioritet
      - Na koga se primjenjuje (filter uvjeti, tip primatelja)
      - Koje email adrese kreira (predlosci adresa)
    Output se sprema u CSV (delimiter ";") koji se moze otvoriti u Excelu ili Wordu.

.PARAMETER OutputPath
    Putanja do izlaznog CSV fajla. Default: .\EmailAddressPolicies_<datum>.csv

.PARAMETER IncludeRecipientCount
    Ako se navede, skripta ce prebrojati korisnike na koje se policy primjenjuje.
    Moze biti sporije u velikim okruzenjima.

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

# Provjera je li Exchange Management Shell vec ucitan
if (Get-Command Get-EmailAddressPolicy -ErrorAction SilentlyContinue) {
    $exchangeLoaded = $true
}

# Pokusaj ucitavanja Exchange snap-ina (Exchange 2010/2013/2016/2019)
if (-not $exchangeLoaded) {
    if (Get-PSSnapin -Registered | Where-Object { $_.Name -eq 'Microsoft.Exchange.Management.PowerShell.SnapIn' }) {
        try {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
            $exchangeLoaded = $true
            Write-Verbose "Exchange snap-in ucitan."
        }
        catch {
            Write-Warning "Nije moguce ucitati Exchange snap-in: $_"
        }
    }
}

# Pokusaj spajanja na Exchange putem implicit remoting (Exchange 2013+)
if (-not $exchangeLoaded) {
    Write-Host "Exchange cmdlets nisu pronadjeni lokalno." -ForegroundColor Yellow
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
        Write-Verbose "Spojeno na Exchange server: $exchangeServer"
    }
    catch {
        Write-Error "Nije moguce spojiti se na Exchange server '$exchangeServer': $_"
        exit 1
    }
}

#endregion

#region --- Pomocne funkcije ---

function Format-TemplateList {
    # Pretvara listu predlozaka email adresa u citljiv string.
    # Oznacava primarnu SMTP adresu s oznakom [PRIMARY].
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
    # Gradi citljiv opis uvjeta primjene iz conditional atributa policy-a.
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
        ConditionalStateOrProvince   = "Drzava/Pokrajina"
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

function ConvertTo-DateString {
    # Pretvara ExDateTime ili [datetime] u formatirani string.
    # ExDateTime ne podrzava ToString("format") pa koristimo -f operator.
    param($Value)

    if (-not $Value) { return "Nikad" }
    try {
        return "{0:yyyy-MM-dd HH:mm}" -f [datetime]$Value
    }
    catch {
        return $Value.ToString()
    }
}

#endregion

#region --- Dohvacanje i obrada Email Address Policies ---

Write-Host "`nDohvacam Exchange organizaciju i Email Address Policies..." -ForegroundColor Cyan

try {
    $orgName = (Get-OrganizationConfig -ErrorAction Stop).Name
}
catch {
    $orgName = "N/A"
}

try {
    $policies = Get-EmailAddressPolicy -ErrorAction Stop | Sort-Object Priority
}
catch {
    Write-Error "Greska pri dohvacanju Email Address Policies: $_"
    exit 1
}

if (-not $policies) {
    Write-Warning "Nisu pronadjene Email Address Policies."
    exit 0
}

Write-Host "Pronadjeno $($policies.Count) polic(y/ies). Obradjujem..." -ForegroundColor Cyan

$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($policy in $policies) {

    # --- Tip filtera i uvjeti primjene ---
    $filterType  = $policy.RecipientFilterType   # Precanned ili Custom
    $ldapFilter  = if ($policy.RecipientFilter) { $policy.RecipientFilter.ToString() } else { "N/A" }

    # Citljivi uvjeti (za Precanned filtere)
    $conditionsFriendly = Format-ConditionalAttributes -Policy $policy

    if ($filterType -eq 'Precanned' -and -not $conditionsFriendly) {
        $conditionsFriendly = "Svi primatelji (nema dodatnih uvjeta)"
    }
    elseif ($filterType -eq 'Custom') {
        $conditionsFriendly = "Prilagodjeni filter (vidi LDAP Filter stupac)"
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
            $recipientCount = "Greska pri dohvacanju"
        }
    }

    # --- Email adresni predlosci ---
    $templates = Format-TemplateList -Templates $policy.EnabledEmailAddressTemplates

    # --- Gradi red za CSV ---
    $row = [PSCustomObject]@{
        "Prioritet"              = $policy.Priority
        "Naziv Policy-a"         = $policy.Name
        "Tip Filtera"            = $filterType
        "Uvjeti Primjene"        = $conditionsFriendly
        "LDAP Filter (Custom)"   = if ($filterType -eq 'Custom') { $ldapFilter } else { "" }
        "Predlosci Email Adresa" = $templates
        "Broj Primatelja"        = $recipientCount
        "Kreirano"               = ConvertTo-DateString -Value $policy.WhenCreated
        "Zadnja Izmjena"         = ConvertTo-DateString -Value $policy.WhenChanged
        "Exchange Organizacija"  = $orgName
    }

    $results.Add($row)

    Write-Verbose "Obradjeno: $($policy.Name) [Prioritet: $($policy.Priority)]"
}

#endregion

#region --- Export u CSV ---

try {
    $outputDir = Split-Path -Path $OutputPath -Parent
    if ($outputDir -and -not (Test-Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    }

    $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    Write-Host "`n[OK] Export uspjesan: $((Resolve-Path $OutputPath).Path)" -ForegroundColor Green
    Write-Host "     Broj redova: $($results.Count)" -ForegroundColor Green
    Write-Host "`nUputa za import u Excel:" -ForegroundColor Yellow
    Write-Host "  Podaci > Iz teksta/CSV > odaberi fajl > Delimiter: Tocka-zarez (;)" -ForegroundColor Yellow
}
catch {
    Write-Error "Greska pri exportu u CSV: $_"
    exit 1
}

#endregion

# Vrati objekte u pipeline za eventualno daljnje procesiranje
$results
