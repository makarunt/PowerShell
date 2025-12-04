# Add-SecondaryEmailAddress.ps1

## Opis

PowerShell skripta za automatsko dodavanje dodatnih email adresa na on-premises Exchange mailboxovima koji imaju isključenu automatsku primjenu Email Address Policy-a.

## Funkcionalnost

Skripta:
1. Pronalazi sve mailboxove gdje je `EmailAddressPolicyEnabled = $false`
2. Za svaki takav mailbox dodaje dodatnu email adresu u formatu: `alias@<domena>`
3. Preskače mailboxove koji već imaju tu adresu
4. Generira detaljan izvještaj u CSV formatu
5. Prikazuje real-time napredak i statistiku

## Preduvjeti

- **Exchange Management Shell** mora biti instaliran i učitan
- **Administratorske privilegije** na Exchange serveru
- **PowerShell 5.1** ili noviji

## Pokretanje skripte

### 1. Otvorite Exchange Management Shell

Pokrenite Exchange Management Shell kao administrator.

### 2. Osnovni primjer

```powershell
# Navigirajte do direktorija sa skriptom
cd C:\Scripts

# Pokrenite skriptu sa obaveznim parametrom domene
.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@company.mail.onmicrosoft.com"
```

### 3. Test režim (WhatIf)

Preporučeno za prvi run - pokazuje što bi skripta napravila bez stvarnih izmjena:

```powershell
.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@company.mail.onmicrosoft.com" -WhatIf
```

### 4. Potvrda za svaki mailbox

```powershell
.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@company.mail.onmicrosoft.com" -Confirm
```

### 5. Različite domene

```powershell
# Primjer 1: Standard onmicrosoft.com domena
.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@tenant.onmicrosoft.com"

# Primjer 2: Custom domena
.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@custom-domain.com"

# Primjer 3: Mail subdomena
.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@mail.company.com"
```

## Parametri

| Parametar | Tip | Obavezan | Opis |
|-----------|-----|----------|------|
| `DomainSuffix` | String | **DA** | Domain sufiks za nove email adrese (npr. @company.mail.onmicrosoft.com) |
| `SkipExisting` | Switch | Ne | Preskače mailboxove koji već imaju adresu |
| `WhatIf` | Switch | Ne | Prikazuje što bi se promijenilo bez stvarnih izmjena |
| `Confirm` | Switch | Ne | Traži potvrdu prije svake izmjene |

### Validacija DomainSuffix parametra

Parametar mora:
- Započinjati sa `@`
- Biti validan format domene (npr. `@example.com`, `@mail.company.onmicrosoft.com`)
- Sadržavati TLD (top-level domain)

**Validni primjeri:**
- ✅ `@company.mail.onmicrosoft.com`
- ✅ `@tenant.onmicrosoft.com`
- ✅ `@custom-domain.com`
- ✅ `@mail.example.co.uk`

**Nevalidni primjeri:**
- ❌ `company.com` (nema @)
- ❌ `@company` (nema TLD)
- ❌ `example.com` (nema @)

## Izlazni podaci

### Konzolni prikaz

Skripta prikazuje:
- ✅ **Success** - zeleno (uspješno dodano)
- ⚠️ **Skip** - žuto (već postoji)
- ❌ **Error** - crveno (greška)

### CSV izvještaj

Generira se CSV izvještaj sa sljedećim kolonama:
- `DisplayName` - Ime mailboxa
- `Alias` - Exchange alias
- `PrimaryEmail` - Primarna email adresa
- `NewEmailAddress` - Novodana email adresa
- `Status` - Status operacije (Success/Skipped/Error)
- `Timestamp` - Vrijeme izvršenja

**Lokacija izvještaja:** `EmailAddressReport_YYYYMMDD_HHMMSS.csv`

## Primjeri output-a

```
Starting email address addition process...
Domain suffix: @company.mail.onmicrosoft.com
--------------------------------------------------------------------------------
Retrieving mailboxes with Email Address Policy disabled...
Found 15 mailbox(es) with Email Address Policy disabled.
--------------------------------------------------------------------------------
[1/15] Processing: John Doe (john.doe@company.com)
  [SUCCESS] Added: smtp:jdoe@company.mail.onmicrosoft.com

[2/15] Processing: Jane Smith (jane.smith@company.com)
  [SKIP] Email address already exists: smtp:jsmith@company.mail.onmicrosoft.com

[3/15] Processing: Mike Johnson (mike.johnson@company.com)
  [SUCCESS] Added: smtp:mjohnson@company.mail.onmicrosoft.com

--------------------------------------------------------------------------------
SUMMARY
--------------------------------------------------------------------------------
Total mailboxes found:     15
Successfully processed:    12
Skipped (already exists):  2
Errors:                    1
--------------------------------------------------------------------------------
Report exported to: C:\Scripts\EmailAddressReport_20251204_143022.csv
Script completed.
```

## Sigurnosne provjere

Skripta uključuje:
- ✅ Provjeru da li je Exchange Management Shell učitan
- ✅ Validaciju formata domene
- ✅ Error handling za svaki mailbox
- ✅ Provjeru postojećih email adresa (ne duplicira)
- ✅ Progress bar za praćenje napretka
- ✅ Detaljno logiranje svih operacija
- ✅ Support za `-WhatIf` i `-Confirm` parametere

## Troubleshooting

### Problem: "Exchange Management Shell is not loaded"

**Rješenje:**
```powershell
# Učitajte Exchange snap-in
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn

# Ili pokrenite direktno Exchange Management Shell
```

### Problem: "Access Denied"

**Rješenje:**
- Pokrenite PowerShell kao administrator
- Provjerite da imate potrebne Exchange RBAC privilegije
- Minimalno potrebna rola: `Mail Recipients` role

### Problem: Execution Policy

**Rješenje:**
```powershell
# Privremeno omogućite izvršavanje skripti
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Ili
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Problem: "Cannot validate argument on parameter 'DomainSuffix'"

**Rješenje:**
- Provjerite da domena započinje sa `@`
- Provjerite da je format validan (npr. `@company.com`)
- Primjer: `.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@company.mail.onmicrosoft.com"`

## Best Practices

1. **Uvijek prvo pokrenite sa `-WhatIf`**
   ```powershell
   .\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@company.mail.onmicrosoft.com" -WhatIf
   ```

2. **Napravite backup prije masovnih izmjena**
   ```powershell
   Get-Mailbox | Export-Clixml MailboxBackup.xml
   ```

3. **Testirajte na malom uzorku**
   - Modificirajte skriptu da procesira samo nekoliko mailboxova
   - Provjerite rezultate
   - Zatim pokrenite na svim mailboxovima

4. **Provjerite CSV izvještaj**
   - Uvijek pregledajte generirani CSV izvještaj
   - Provjerite status za svaki mailbox

5. **Dokumentirajte domenu**
   - Zapišite koju domenu ste koristili
   - Spremite CSV izvještaj za buduću referencu

## Napomene

- Skripta **ne mijenja** primarnu email adresu mailboxa
- Skripta **dodaje** novu sekundarnu email adresu
- Email adresa se dodaje kao `smtp:` (mala slova), ne kao `SMTP:` (primarna)
- Mailboxovi sa već postojećom adresom se preskače
- **DomainSuffix parametar je obavezan** - skripta neće raditi bez njega

## Rollback procedura

Ako trebate ukloniti dodane email adrese:

```powershell
# Definirajte domenu koju ste koristili
$domainToRemove = "@company.mail.onmicrosoft.com"

# Dohvatite mailboxove
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
    $_.EmailAddressPolicyEnabled -eq $false
}

# Uklonite specifične adrese
foreach ($mailbox in $mailboxes) {
    $addressToRemove = "smtp:$($mailbox.Alias)$domainToRemove"

    if ($mailbox.EmailAddresses -contains $addressToRemove) {
        Set-Mailbox -Identity $mailbox.Identity `
                    -EmailAddresses @{Remove=$addressToRemove}
        Write-Host "Removed: $addressToRemove from $($mailbox.DisplayName)"
    }
}
```

## Primjer full workflow-a

```powershell
# 1. Otvorite Exchange Management Shell kao administrator

# 2. Navigirajte do foldera sa skriptom
cd C:\Scripts

# 3. Testirajte sa WhatIf
.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@company.mail.onmicrosoft.com" -WhatIf

# 4. Pregledajte output i odlučite da li nastaviti

# 5. Pokrenite stvarno izvršavanje
.\Add-SecondaryEmailAddress.ps1 -DomainSuffix "@company.mail.onmicrosoft.com"

# 6. Pregledajte CSV izvještaj
Import-Csv .\EmailAddressReport_<timestamp>.csv | Out-GridView

# 7. Provjerite nekoliko mailboxova
Get-Mailbox "John Doe" | Select-Object DisplayName, EmailAddresses

# 8. Ako je sve OK, završeno!
# 9. Ako trebate rollback, koristite proceduru gore
```

## Sigurnosne preporuke

- ⚠️ **Nikad** ne pokrećite skriptu u production okruženju bez `-WhatIf` testa
- ⚠️ **Uvijek** napravite backup prije masovnih promjena
- ⚠️ **Provjerite** da li imate ispravnu domenu prije pokretanja
- ⚠️ **Čuvajte** CSV izvještaje za compliance i audit trail
- ⚠️ **Testirajte** na test mailboxovima prije production-a

## Kontakt i podrška

Za pitanja ili probleme:
- Kreirajte GitHub Issue
- Kontaktirajte Exchange administratora

## Licenca

MIT License

---

**Verzija:** 2.0
**Datum:** 2025-12-04
**Autor:** PowerShell Script
