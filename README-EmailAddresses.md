# Add-FortenovaEmailAddresses.ps1

## Opis

PowerShell skripta za automatsko dodavanje dodatnih email adresa na on-premises Exchange mailboxovima koji imaju isključenu automatsku primjenu Email Address Policy-a.

## Funkcionalnost

Skripta:
1. Pronalazi sve mailboxove gdje je `EmailAddressPolicyEnabled = $false`
2. Za svaki takav mailbox dodaje dodatnu email adresu u formatu: `alias@fortenova.mail.onmicrosoft.com`
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

# Pokrenite skriptu
.\Add-FortenovaEmailAddresses.ps1
```

### 3. Test režim (WhatIf)

Preporučeno za prvi run - pokazuje što bi skripta napravila bez stvarnih izmjena:

```powershell
.\Add-FortenovaEmailAddresses.ps1 -WhatIf
```

### 4. Potvrda za svaki mailbox

```powershell
.\Add-FortenovaEmailAddresses.ps1 -Confirm
```

### 5. Korištenje custom domene

```powershell
.\Add-FortenovaEmailAddresses.ps1 -DomainSuffix "@custom.onmicrosoft.com"
```

## Parametri

| Parametar | Tip | Opis | Default |
|-----------|-----|------|---------|
| `DomainSuffix` | String | Domain sufiks za nove email adrese | `@fortenova.mail.onmicrosoft.com` |
| `SkipExisting` | Switch | Preskače mailboxove koji već imaju adresu | `$false` |
| `WhatIf` | Switch | Prikazuje što bi se promijenilo bez stvarnih izmjena | - |
| `Confirm` | Switch | Traži potvrdu prije svake izmjene | - |

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
Domain suffix: @fortenova.mail.onmicrosoft.com
--------------------------------------------------------------------------------
Retrieving mailboxes with Email Address Policy disabled...
Found 15 mailbox(es) with Email Address Policy disabled.
--------------------------------------------------------------------------------
[1/15] Processing: John Doe (john.doe@company.com)
  [SUCCESS] Added: smtp:jdoe@fortenova.mail.onmicrosoft.com

[2/15] Processing: Jane Smith (jane.smith@company.com)
  [SKIP] Email address already exists: smtp:jsmith@fortenova.mail.onmicrosoft.com

[3/15] Processing: Mike Johnson (mike.johnson@company.com)
  [SUCCESS] Added: smtp:mjohnson@fortenova.mail.onmicrosoft.com

--------------------------------------------------------------------------------
SUMMARY
--------------------------------------------------------------------------------
Total mailboxes found:     15
Successfully processed:    12
Skipped (already exists):  2
Errors:                    1
--------------------------------------------------------------------------------
Report exported to: C:\Scripts\EmailAddressReport_20251203_143022.csv
Script completed.
```

## Sigurnosne provjere

Skripta uključuje:
- ✅ Provjeru da li je Exchange Management Shell učitan
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

## Best Practices

1. **Uvijek prvo pokrenite sa `-WhatIf`**
   ```powershell
   .\Add-FortenovaEmailAddresses.ps1 -WhatIf
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

## Napomene

- Skripta **ne mijenja** primarnu email adresu mailboxa
- Skripta **dodaje** novu sekundarnu email adresu
- Email adresa se dodaje kao `smtp:` (mala slova), ne kao `SMTP:` (primarna)
- Mailboxovi sa već postojećom adresom se preskače

## Rollback procedura

Ako trebate ukloniti dodane email adrese:

```powershell
# Dohvatite mailboxove
$mailboxes = Get-Mailbox -ResultSize Unlimited | Where-Object {
    $_.EmailAddressPolicyEnabled -eq $false
}

# Uklonite specifične adrese
foreach ($mailbox in $mailboxes) {
    $addressToRemove = "smtp:$($mailbox.Alias)@fortenova.mail.onmicrosoft.com"

    if ($mailbox.EmailAddresses -contains $addressToRemove) {
        Set-Mailbox -Identity $mailbox.Identity `
                    -EmailAddresses @{Remove=$addressToRemove}
        Write-Host "Removed: $addressToRemove from $($mailbox.DisplayName)"
    }
}
```

## Kontakt i podrška

Za pitanja ili probleme:
- Kreirajte GitHub Issue
- Kontaktirajte Exchange administratora

## Licenca

MIT License

---

**Verzija:** 1.0
**Datum:** 2025-12-03
**Autor:** PowerShell Script
