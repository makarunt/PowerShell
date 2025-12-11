# Exchange Online Mailbox Migration Scripts

Ova kolekcija PowerShell skripti omogućava jednostavno praćenje i upravljanje remote mailbox migracijama sa on-premises Exchange servera na Exchange Online.

## 📋 Sadržaj

- [Get-MailboxMigrationDuration.ps1](#get-mailboxmigrationdurationps1) - Praćenje trajanja migracija do 95%
- [Test-MigrationReadyForCompletion.ps1](#test-migrationreadyforcompletionps1) - Brza provjera spremnosti za finalizaciju
- [Send-MigrationStatusReport.ps1](#send-migrationstatusreportps1) - Automatsko slanje email izvještaja

---

## Get-MailboxMigrationDuration.ps1

### 📖 Opis

Ova skripta omogućava detaljno praćenje vremena trajanja mailbox migracija i prikazuje **točno vrijeme potrebno da migration batch dođe do 95%** - što je kritična faza kada se može pokrenuti finalizacija migracije.

### ✨ Ključne funkcionalnosti

- ✅ **Praćenje vremena do 95%** - Izračunava točno vrijeme potrebno da migracija dostigne 95% (faza spremna za finalizaciju)
- 📊 **Detaljne statistike** - Prikazuje sve relevantne podatke o migraciji (bytes transferred, items, progress, itd.)
- ⏱️ **Procjena preostalog vremena** - Za migracije u toku procjenjuje kada će dostići 95%
- 📁 **CSV export** - Mogućnost izvoza svih podataka u CSV format
- 🎨 **Čitljiv output** - Koristi boje i emoji za lakše prepoznavanje statusa
- 🔍 **Filtriranje po batch-u** - Može pratiti specifični migration batch ili sve odjednom

### 🔧 Preduvjeti

1. **ExchangeOnlineManagement modul**
   ```powershell
   Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
   ```

2. **Exchange Online konekcija**
   ```powershell
   Connect-ExchangeOnline -UserPrincipalName admin@contoso.com
   ```

3. **Permisije**
   - Organization Management ili
   - Migration Administrator role

### 📝 Parametri

| Parametar | Tip | Obavezan | Opis |
|-----------|-----|----------|------|
| `BatchName` | String | Ne | Naziv specifičnog migration batcha |
| `IncludeCompleted` | Switch | Ne | Uključuje completed batch-eve u rezultate |
| `ExportToCsv` | Switch | Ne | Izvozi rezultate u CSV datoteku |
| `CsvPath` | String | Ne | Putanja za CSV export (default: trenutni direktorij + timestamp) |
| `ShowDetailedStats` | Switch | Ne | Prikazuje detaljne statistike za svaki mailbox |

### 💡 Primjeri upotrebe

#### Primjer 1: Praćenje svih aktivnih migracija

```powershell
.\Get-MailboxMigrationDuration.ps1
```

**Output:**
```
═══════════════════════════════════════════════════════════════
   Analiza trajanja mailbox migracija na Exchange Online
═══════════════════════════════════════════════════════════════

Pronađeno 2 migration batch(eva)

─────────────────────────────────────────────────────────────────
Batch: MigrationBatch-Finance-2024
─────────────────────────────────────────────────────────────────
  Status: Syncing
  Total Mailboxes: 15
  Synced: 12
  Active: 3
  Failed: 0
  Created: 08.12.2024 10:30:15

  Analiziram 15 mailbox(eva)...

  📧 john.doe@contoso.com
     Status: Synced
     Progress: 98%
     ⏱️  Vrijeme do 95%: 2 dana, 14 sati, 35 minuta
     ⌚ Ukupno trajanje: 2 dana, 16 sati, 22 minuta
     ✓ SPREMAN ZA FINALIZACIJU!

  📧 jane.smith@contoso.com
     Status: Syncing
     Progress: 67%
     ⌚ Ukupno trajanje: 1 dana, 8 sati, 15 minuta
     ⏳ Procjena do 95%: 10.12.2024 18:45
```

#### Primjer 2: Praćenje specifičnog batcha sa detaljnim statistikama

```powershell
.\Get-MailboxMigrationDuration.ps1 -BatchName "MigrationBatch-Finance-2024" -ShowDetailedStats
```

**Output uključuje dodatne detalje:**
```
  📧 john.doe@contoso.com
     Status: Synced
     Progress: 98%
     ⏱️  Vrijeme do 95%: 2 dana, 14 sati, 35 minuta
     ⌚ Ukupno trajanje: 2 dana, 16 sati, 22 minuta
     ✓ SPREMAN ZA FINALIZACIJU!
     ├─ Bytes Transferred: 25.8 GB (27,730,123,456 bytes)
     ├─ Total Mailbox Size: 26.3 GB (28,240,987,654 bytes)
     ├─ Items Transferred: 45,678
     ├─ Total Items: 46,234
     ├─ Sync Started: 08.12.2024 10:35:22
     └─ Last Synced: 11.12.2024 02:58:11
```

#### Primjer 3: Export svih migracija u CSV

```powershell
.\Get-MailboxMigrationDuration.ps1 -IncludeCompleted -ExportToCsv
```

Kreira CSV datoteku: `MigrationDuration_20241211_143522.csv`

#### Primjer 4: Automatsko praćenje sa custom CSV putanjom

```powershell
.\Get-MailboxMigrationDuration.ps1 -BatchName "Batch-IT-Department" -ExportToCsv -CsvPath "C:\Reports\IT_Migration.csv"
```

#### Primjer 5: Praćenje samo aktivnih migracija bez completed

```powershell
# Default ponašanje - ne prikazuje completed batch-eve
.\Get-MailboxMigrationDuration.ps1
```

### 📊 Razumijevanje outputa

#### Statusne boje

- 🟢 **Zeleno (Synced)** - Migracija završena, sinkronizirano
- 🔵 **Cyan (Syncing)** - Migracija u toku
- 🟡 **Žuto (Queued)** - Čeka na početak
- 🔴 **Crveno (Failed)** - Greška u migraciji

#### Vrijeme do 95%

Skripta izračunava vrijeme kada je mailbox dostigao (ili će dostići) **95% completion**. Ovo je važno jer:

1. **Na 95%**, migration batch je spreman za finalizaciju
2. **Finalizacija se pokreće zasebno** komandom `Complete-MigrationBatch`
3. **Finalizacija** je brza operacija koja prebacuje mailbox sa on-prem na EXO

#### Procjena preostalog vremena

Za migracije u toku (<95%), skripta:
- Izračunava prosječnu brzinu migracije
- Procjenjuje kada će dostići 95%
- Prikazuje estimirano vrijeme završetka

**Napomena:** Procjena se bazira na linearnoj progresiji i može varirati ovisno o veličini podataka i mrežnim uvjetima.

### 🔄 Tipični workflow za finalizaciju

1. **Pokrenite skriptu** da provjerite status:
   ```powershell
   .\Get-MailboxMigrationDuration.ps1 -BatchName "YourBatch"
   ```

2. **Pričekajte** da mailboxevi dostignu 95% (skripta će prikazati "SPREMAN ZA FINALIZACIJU!")

3. **Finalizirajte migration batch**:
   ```powershell
   Complete-MigrationBatch -Identity "YourBatch"
   ```

4. **Praćenje finalizacije**:
   ```powershell
   Get-MigrationBatch -Identity "YourBatch" | Format-List Status,LastSyncedDate
   ```

### 📁 CSV Export struktura

Izvezena CSV datoteka sadrži sljedeće kolone:

| Kolona | Opis |
|--------|------|
| Identity | Email adresa mailboxa |
| BatchId | Naziv migration batcha |
| Status | Trenutni status (Synced, Syncing, Failed, itd.) |
| PercentageComplete | Postotak završetka |
| IsReadyForFinalization | True/False - spreman za Complete-MigrationBatch |
| TimeTo95PercentFormatted | Formatirano vrijeme do 95% |
| TotalDurationFormatted | Ukupno trajanje migracije |
| StartDate | Datum početka migracije |
| InitialSyncDateTime | Datum/vrijeme početka sync-a |
| LastSyncedDateTime | Datum/vrijeme zadnjeg sync-a |
| EstimatedCompletion | Procijenjeno vrijeme dostizanja 95% |
| BytesTransferred | Broj byte-ova transferiranih |
| TotalMailboxSize | Ukupna veličina mailboxa |
| ItemsTransferred | Broj transferiranih itema |
| TotalItemsInMailbox | Ukupan broj itema u mailboxu |
| Message | Statusna poruka |

### ⚠️ Napomene i best practices

1. **Konekcija**
   - Obavezno se spojite na Exchange Online prije pokretanja skripte
   - Koristite `Connect-ExchangeOnline` komandu

2. **Performanse**
   - Za veće migracije (100+ mailboxeva), izvršavanje može potrajati
   - Koristite `-BatchName` parametar za praćenje specifičnog batcha

3. **Praćenje u realnom vremenu**
   - Pokrenite skriptu periodično da pratite napredak
   - Kombinirajte sa Task Schedulerom za automatsko izvještavanje

4. **95% threshold**
   - Ovo je Microsoft preporučeni threshold za finalizaciju
   - Finalizacija na <95% može rezultirati dužim cutover-om

5. **Finalizacija**
   - Finalizaciju **uvijek** planirajte za vrijeme minimalne aktivnosti (noć, vikend)
   - Tijekom finalizacije korisnici ne mogu pristupiti mailboxu (kratko vrijeme)

### 🐛 Troubleshooting

#### Greška: "Nije uspostavljena konekcija sa Exchange Online"

**Rješenje:**
```powershell
Connect-ExchangeOnline -UserPrincipalName admin@contoso.com
```

#### Greška: "Access Denied" ili permisije

**Rješenje:** Provjerite da imate jednu od sljedećih rola:
- Organization Management
- Migration Administrator

```powershell
# Provjera trenutnih rola
Get-ManagementRoleAssignment -RoleAssignee "your.admin@contoso.com" |
    Format-Table Role,RoleAssigneeType -AutoSize
```

#### Prazan output ili nema migration batch-eva

**Mogući uzroci:**
1. Nisu kreirani migration batch-evi
2. Svi batch-evi su completed (koristite `-IncludeCompleted`)
3. Batch Name je pogrešan

**Provjera:**
```powershell
# Lista svih batch-eva
Get-MigrationBatch | Format-Table Identity,Status,TotalCount -AutoSize
```

#### Statistike nisu dostupne

Ako `Get-MigrationUserStatistics` ne vraća podatke:
```powershell
# Ručna provjera
Get-MigrationUser -BatchId "YourBatch" | Select-Object Identity,Status
Get-MigrationUserStatistics -Identity "user@contoso.com"
```

### 📚 Dodatne korisne komande

#### Kreiranje migration batcha

```powershell
New-MigrationBatch -Name "Batch-Finance-2024" `
    -SourceEndpoint "OnPremEndpoint" `
    -CSVData ([System.IO.File]::ReadAllBytes("C:\migration.csv")) `
    -AutoStart
```

#### Pokretanje migracije

```powershell
Start-MigrationBatch -Identity "Batch-Finance-2024"
```

#### Finalizacija nakon 95%

```powershell
Complete-MigrationBatch -Identity "Batch-Finance-2024" -Confirm:$false
```

#### Uklanjanje completed batcha

```powershell
Remove-MigrationBatch -Identity "Batch-Finance-2024" -Confirm:$false
```

#### Zaustavljanje migracije

```powershell
Stop-MigrationBatch -Identity "Batch-Finance-2024"
```

### 🔗 Korisni linkovi

- [Microsoft: Migrate mailboxes to Exchange Online](https://docs.microsoft.com/en-us/exchange/mailbox-migration/mailbox-migration)
- [Get-MigrationBatch documentation](https://docs.microsoft.com/en-us/powershell/module/exchange/get-migrationbatch)
- [Complete-MigrationBatch documentation](https://docs.microsoft.com/en-us/powershell/module/exchange/complete-migrationbatch)
- [Exchange Online PowerShell](https://docs.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)

### 📧 Support

Za pitanja i probleme, kreirajte issue ili kontaktirajte administratora.

---

## Test-MigrationReadyForCompletion.ps1

### 📖 Opis

Brza i jednostavna skripta za provjeru koji migration batch-evi i mailboxevi su dostigli **95% ili više** i spremni su za finalizaciju. Idealna za brzu provjeru prije pokretanja `Complete-MigrationBatch` komande.

### ✨ Ključne funkcionalnosti

- ⚡ **Brza provjera** - Prikazuje samo mailboxeve spremne za finalizaciju (≥95%)
- 🎯 **Fokus na akciju** - Direktno prikazuje što je spremno za Complete-MigrationBatch
- 🚀 **AutoComplete opcija** - Može automatski pokrenuti finalizaciju (OPREZ!)
- 📋 **Sažeti prikaz** - Jednostavan i pregledan output

### 📝 Parametri

| Parametar | Tip | Obavezan | Opis |
|-----------|-----|----------|------|
| `BatchName` | String | Ne | Naziv specifičnog migration batcha za provjeru |
| `AutoComplete` | Switch | Ne | Automatski pokreće Complete-MigrationBatch (BEZ POTVRDE!) |

### 💡 Primjeri upotrebe

#### Primjer 1: Brza provjera svih batch-eva

```powershell
.\Test-MigrationReadyForCompletion.ps1
```

**Output:**
```
╔════════════════════════════════════════════════════════════╗
║  Provjera spremnosti za finalizaciju migracije (≥95%)     ║
╚════════════════════════════════════════════════════════════╝

✅ BATCH SPREMAN: MigrationBatch-Finance-2024
   📊 Mailboxeva: 15/15 (100%)
   📧 Mailboxevi spremni za finalizaciju:
      • john.doe@contoso.com - 98% (trajanje: 2d 14h 35m)
      • jane.smith@contoso.com - 96% (trajanje: 2d 8h 12m)
      • bob.johnson@contoso.com - 100% (trajanje: 1d 22h 5m)

   ▶️  Komanda za finalizaciju:
      Complete-MigrationBatch -Identity 'MigrationBatch-Finance-2024'
```

#### Primjer 2: Provjera specifičnog batcha

```powershell
.\Test-MigrationReadyForCompletion.ps1 -BatchName "Batch-IT-Department"
```

#### Primjer 3: Automatska finalizacija (OPREZ!)

```powershell
.\Test-MigrationReadyForCompletion.ps1 -AutoComplete
```

**UPOZORENJE:** Ovo će automatski pokrenuti `Complete-MigrationBatch` za sve batch-eve gdje su SVI mailboxevi ≥95%, bez dodatne potvrde!

### ⚠️ Napomene

- **AutoComplete** - Koristite samo ako ste sigurni! Ne traži potvrdu.
- Prikazuje samo batch-eve koji imaju mailboxeve ≥95%
- Batch je "potpuno spreman" samo ako su SVI mailboxevi u njemu ≥95%

---

## Send-MigrationStatusReport.ps1

### 📖 Opis

Automatski generiše i šalje HTML email izvještaj sa trenutnim statusom svih migration batch-eva. Idealno za scheduliranje sa Windows Task Schedulerom za redovno izvještavanje IT tima ili managementa.

### ✨ Ključne funkcionalnosti

- 📧 **HTML Email izvještaj** - Profesionalno formatiran izvještaj sa bojama i grafičkim prikazom
- 📊 **Vizualni progress bar** - Grafički prikaz napretka svake migracije
- 📅 **Scheduling ready** - Može se pokrenuti automatski putem Task Schedulera
- 🎯 **Filtriranje** - Opcija za slanje samo mailboxeva spremnih za finalizaciju
- 💾 **Lokalna sačuva** - Ako email ne može biti poslan, sačuva HTML lokalno

### 📝 Parametri

| Parametar | Tip | Obavezan | Opis |
|-----------|-----|----------|------|
| `To` | String[] | Da | Email adresa(e) primaoca (odvojene zarezom) |
| `From` | String | Ne | Email adresa pošiljaoca (default: noreply@tenant.onmicrosoft.com) |
| `BatchName` | String | Ne | Naziv specifičnog migration batcha |
| `SmtpServer` | String | Ne | SMTP server (default: smtp.office365.com) |
| `IncludeOnlyReady` | Switch | Ne | Uključuje samo mailboxeve spremne za finalizaciju |

### 💡 Primjeri upotrebe

#### Primjer 1: Pošalji izvještaj IT timu

```powershell
.\Send-MigrationStatusReport.ps1 -To "it-team@contoso.com"
```

#### Primjer 2: Izvještaj samo o spremnim mailboxevima

```powershell
.\Send-MigrationStatusReport.ps1 -To "admin@contoso.com" -IncludeOnlyReady
```

#### Primjer 3: Specifični batch sa custom From adresom

```powershell
.\Send-MigrationStatusReport.ps1 `
    -BatchName "Batch-Finance-2024" `
    -To "finance-team@contoso.com,it-admin@contoso.com" `
    -From "migration-reports@contoso.com"
```

#### Primjer 4: Scheduliranje sa Task Schedulerom

**PowerShell komanda za kreiranje scheduled task-a:**

```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument @"
-NoProfile -ExecutionPolicy Bypass -Command "& {
    Connect-ExchangeOnline -CertificateThumbprint 'YOUR_CERT_THUMBPRINT' -AppId 'YOUR_APP_ID' -Organization 'contoso.onmicrosoft.com'
    C:\Scripts\Send-MigrationStatusReport.ps1 -To 'it-team@contoso.com'
}"
"@

$trigger = New-ScheduledTaskTrigger -Daily -At "08:00AM"

Register-ScheduledTask -TaskName "Daily Migration Status Report" `
    -Action $action `
    -Trigger $trigger `
    -Description "Šalje dnevni izvještaj o migration statusu"
```

### 📧 Email izvještaj sadrži

- **Sažetak**: Ukupno batch-eva, mailboxeva, koliko je spremno, u toku, failed
- **Detalji po batchu**:
  - Status batcha
  - Lista svih mailboxeva
  - Progress bar za svaki mailbox
  - Trajanje migracije
  - Bytes transferred
- **Vizualne boje**: Zeleno (Synced), Plavo (Syncing), Crveno (Failed), Žuto (Queued)

### ⚠️ Napomene

1. **SMTP autentifikacija**:
   - `Send-MailMessage` cmdlet je deprecated ali još uvijek funkcionalan
   - Za produkciju, preporučujem korištenje **Microsoft Graph API** za slanje email-a
   - Možda će biti potrebne SMTP credentials (vidi primjer u skripti)

2. **Alternativa - lokalna sačuva**:
   - Ako email ne može biti poslan, skripta automatski sačuva HTML report lokalno
   - HTML fajl može biti ručno poslan ili pregledan u browseru

3. **Certificate-based authentication za scheduling**:
   - Za automatsko izvršavanje koristite app-based authentication sa certifikatom
   - Izbjegavajte čuvanje passworda u plain text-u

### 🔧 Napredna konfiguracija - Microsoft Graph API

Za produkcijsko slanje email-a putem Graph API-ja:

```powershell
# Instalacija Graph modula
Install-Module Microsoft.Graph -Scope CurrentUser

# Konekcija
Connect-MgGraph -Scopes "Mail.Send"

# Slanje (primjer)
$mailParams = @{
    Message = @{
        Subject = "Migration Status Report"
        Body = @{
            ContentType = "HTML"
            Content = $htmlReport
        }
        ToRecipients = @(
            @{ EmailAddress = @{ Address = "recipient@contoso.com" } }
        )
    }
}

Send-MgUserMail -UserId "sender@contoso.com" -BodyParameter $mailParams
```

---

## 🔄 Kompletni Workflow

Kombinacija svih skripti za optimalan migration workflow:

### 1. **Početak migracije**
```powershell
# Kreirajte i pokrenite migration batch
New-MigrationBatch -Name "Batch-Q1-2025" -SourceEndpoint "OnPrem" -CSVData $csv -AutoStart
```

### 2. **Praćenje napretka** (periodično)
```powershell
# Detaljno praćenje sa izvozom u CSV
.\Get-MailboxMigrationDuration.ps1 -BatchName "Batch-Q1-2025" -ExportToCsv
```

### 3. **Brza provjera spremnosti**
```powershell
# Svaki dan provjerite ko je spreman
.\Test-MigrationReadyForCompletion.ps1
```

### 4. **Automatski dnevni izvještaji**
```powershell
# Postavite scheduled task da šalje izvještaje svakog jutra
# (vidi primjer sa Task Schedulerom gore)
```

### 5. **Finalizacija**
```powershell
# Kada su svi mailboxevi ≥95%, finalizirajte
Complete-MigrationBatch -Identity "Batch-Q1-2025"
```

### 6. **Post-finalizacija provjera**
```powershell
# Provjerite status nakon finalizacije
Get-MigrationBatch -Identity "Batch-Q1-2025" | Format-List
Get-MigrationUser -BatchId "Batch-Q1-2025" | Where-Object {$_.Status -ne 'Completed'}
```

---

## 📦 Instalacija

### Brza instalacija

```powershell
# 1. Klonirajte repository ili preuzmite skripte
git clone <repository-url>
cd PowerShell/scripts/ExchangeOnline

# 2. Instalirajte potrebne module
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber

# 3. Spojite se na Exchange Online
Connect-ExchangeOnline -UserPrincipalName admin@contoso.com

# 4. Pokrenite skripte
.\Get-MailboxMigrationDuration.ps1
```

### Provjera permisija

```powershell
# Provjerite da imate potrebne role
Get-ManagementRoleAssignment -RoleAssignee "your.admin@contoso.com" |
    Where-Object {$_.Role -like "*Migration*" -or $_.Role -like "*Organization Management*"} |
    Format-Table Role, RoleAssigneeType -AutoSize
```

---

## 🆘 Troubleshooting - Zajedničke greške

### 1. "ExchangeOnlineManagement modul nije pronađen"

```powershell
Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
Import-Module ExchangeOnlineManagement
```

### 2. "Connect-ExchangeOnline nije prepoznat"

```powershell
# Reinstalirajte modul
Uninstall-Module ExchangeOnlineManagement -Force
Install-Module ExchangeOnlineManagement -Force
```

### 3. Email se ne šalje

**Opcija A: Koristite Graph API** (preporučeno)
**Opcija B: Omogućite SMTP AUTH za mailbox**
```powershell
Set-CASMailbox -Identity "sender@contoso.com" -SmtpClientAuthenticationDisabled $false
```

### 4. "Access Denied" greške

```powershell
# Dodajte potrebnu rolu
New-ManagementRoleAssignment -Role "Migration Administrator" -User "admin@contoso.com"
```

---

## 📚 Dodatni resursi

### Microsoft dokumentacija
- [Exchange Online Migration](https://learn.microsoft.com/en-us/exchange/mailbox-migration/mailbox-migration)
- [Migration Cmdlets](https://learn.microsoft.com/en-us/powershell/module/exchange/?view=exchange-ps#migration)
- [Exchange Online PowerShell](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell)

### Best Practices
- [Migration Performance Best Practices](https://learn.microsoft.com/en-us/exchange/mailbox-migration/office-365-migration-best-practices)
- [Batch Migration Guide](https://learn.microsoft.com/en-us/exchange/mailbox-migration/migrating-imap-mailboxes/batch-migration-to-exchange-online)

---

**Verzija:** 1.0
**Zadnje ažuriranje:** 11.12.2024
**Autor:** PowerShell Migration Scripts
