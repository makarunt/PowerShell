# Invoke-SoaConversion.ps1

Alat za promjenu Source of Authority (SOA) u hibridnom Microsoft 365 okruženju
(on-prem AD + Entra Connect Sync + Exchange Online + Cloud Sync EXO→AD writeback).

Skripta mijenja **samo SOA zastavicu**:

| Kategorija | Promjena | Provjera |
|---|---|---|
| `ExchangeAttributes` | `Set-Mailbox -Identity <ExternalDirectoryObjectId> -IsExchangeCloudManaged $true/$false` | `Get-Mailbox` → `IsExchangeCloudManaged` |
| `User` | `PATCH /users/{id}/onPremisesSyncBehavior` `{ "isCloudManaged": true/false }` | `GET /users/{id}/onPremisesSyncBehavior` |
| `Group` | `PATCH /groups/{id}/onPremisesSyncBehavior` | `GET /groups/{id}/onPremisesSyncBehavior` |
| `Contact` | `PATCH /contacts/{id}/onPremisesSyncBehavior` | `GET /contacts/{id}/onPremisesSyncBehavior` |

Skripta **ne** mijenja tenant-wide postavke (npr. `Set-OrganizationConfig -ExchangeAttributesCloudManagedByDefault`,
`blockCloudObjectTakeoverThroughHardMatchEnabled`), ne dira on-prem AD, Entra Connect ni Cloud Sync, ne mijenja
druge atribute objekata i ne koristi app-only autentifikaciju (ako je postojeća EXO ili Graph konekcija app-only,
skripta staje).

---

## Preduvjeti

| Što | Verzija / napomena |
|---|---|
| PowerShell | **7.2 ili noviji**. Windows PowerShell 5.1 nije podržan i ne treba ga. |
| `ExchangeOnlineManagement` | 3.2.0+ (potreban za `Get-ConnectionInformation`; `Set-Mailbox` mora imati `-IsExchangeCloudManaged`, skripta to provjerava). |
| `Microsoft.Graph.Authentication` | 2.x. Svi Graph pozivi idu kroz `Invoke-MgGraphRequest`. |
| Entra Connect Sync | Za ExchangeAttributes **2.5.190.0+** (starije verzije pokušavaju gurnuti Exchange atribute već prebačenih mailboxova u Entra ID i ne uspijevaju). Za User/Group SOA min. 2.5.76.0, za **Contact SOA 2.5.79.0**. Provjera: `(Get-ADSyncGlobalSettings).Parameters['Microsoft.Synchronize.ServerConfigurationVersion']` |
| Cloud Sync agent | Min. 1.1.1370.0 za User/Group SOA; **1.1.1107.0+** za Exchange attribute writeback (konfiguracija *EXO to AD attribute sync*). Cloud Sync radi uz Connect Sync, ne umjesto njega. |
| Prijava | Interaktivna (delegated). Za `IsExchangeCloudManaged` app-only/certifikat nije podržan. |

Provjera pri pokretanju:
- **PowerShell:** ako je verzija starija od 7.2 (npr. Windows PowerShell 5.1), skripta ispiše potrebnu verziju i
  kako instalirati PowerShell 7 (`winget install --id Microsoft.PowerShell --source winget`), te završi s exit
  kodom 3 bez ikakve promjene i bez izlazne mape. Skriptu treba pokretati iz `pwsh`, ne iz `powershell`.
- **Moduli:** provjeravaju se samo moduli potrebni za odabranu kategoriju. Ako modul nedostaje ili je prestar,
  skripta pita `Install '<modul>' from the PowerShell Gallery for the current user now? [Y/N]`.
  - Na `Y` instalira modul (`Install-Module -Scope CurrentUser -Repository PSGallery -Force`, uz starije
    verzije ako postoje) i nastavlja.
  - Na bilo koji drugi odgovor staje i ispiše naredbu za ručnu instalaciju.
  - Ako je u istoj sesiji već učitana prestara verzija modula, skripta traži novu PowerShell sesiju.
  - Instalacija se radi i kad je skripta pokrenuta s `-WhatIf`, jer je to lokalna promjena koju si upravo
    potvrdio.

Ručna instalacija, ako je radije želiš unaprijed:

```powershell
Install-Module ExchangeOnlineManagement -MinimumVersion 3.2.0 -Scope CurrentUser
```
```powershell
Install-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -Scope CurrentUser
```

Skripta se spaja samo na servise koji su kategoriji potrebni, i to prvo na EXO, a zatim na Graph (redoslijed
smanjuje rizik od sukoba MSAL biblioteka između dva modula). Ako su obje konekcije aktivne, provjerava da su na
istom tenantu.

| Kategorija | EXO | Graph |
|---|---|---|
| ExchangeAttributes | da | ne |
| User | ne | da |
| Group | za Inventory/Convert/Rollback kad `-GroupType` nije `Security` (EXO svojstva mail-enabled grupa u backupu) | da |
| Contact | za Inventory/Convert/Rollback (EXO svojstva kontakta u backupu) | da |

## Uloge i dozvole

| Kategorija | Entra / EXO uloga (minimalno) | Graph delegated scope koji skripta traži |
|---|---|---|
| ExchangeAttributes | Exchange Administrator (Microsoft preporučuje), ili Hybrid Identity Administrator, ili Global Administrator | – |
| User | **Hybrid Administrator** | `User.Read.All`, `User-OnPremisesSyncBehavior.ReadWrite.All` |
| Group | **Hybrid Administrator** (+ EXO čitanje za mail-enabled grupe, npr. View-Only Recipients) | `Group.Read.All`, `GroupMember.Read.All`, `User.Read.All`, `Group-OnPremisesSyncBehavior.ReadWrite.All` |
| Contact | **Hybrid Administrator** (+ EXO čitanje kontakata) | `OrgContact.Read.All`, `Contacts-OnPremisesSyncBehavior.ReadWrite.All` |

- I `GET` i `PATCH` na `onPremisesSyncBehavior` traže `*-OnPremisesSyncBehavior.ReadWrite.All`, pa ga traži i Inventory.
- Prvi put netko s ulogom **Application Administrator** ili **Cloud Application Administrator** mora dati
  consent za ove scopeove aplikaciji *Microsoft Graph Command Line Tools* (Entra admin center →
  Enterprise applications → Microsoft Graph Command Line Tools → Permissions → Grant admin consent).
  Ako scope nije dodijeljen, skripta staje prije bilo kakvog čitanja ili promjene.

## Parametri

```
-Action        Inventory | Convert | Verify | Rollback       (obavezan)
-Category      ExchangeAttributes | User | Group | Contact   (obavezan)
-InputCsv      obavezan za Convert/Verify/Rollback; opcionalan za Inventory
-GroupType     Security | MailEnabledSecurity | Distribution | All   (samo Category=Group, default All)
-OutputFolder  default .\SoaOutput
-MaxObjects    obradi najviše N objekata (prvih N redova CSV-a / prvih N prihvatljivih kod Inventoryja)
-Delimiter     default ';'
-ExpectedState CloudManaged (default) | OnPremManaged   (samo Verify)
-WhatIf / -Confirm
```

`-ExpectedState` nije bio u zahtjevu. Dodan je kako bi se samostalni Verify mogao koristiti i nakon rollbacka.

Za Group, `-GroupType` se primjenjuje u svim akcijama: grupa drugog tipa u CSV-u preskače se s napomenom
`GroupTypeNotSelected`. Microsoft 365 (Unified) grupe nikad nisu prihvatljive.

## Preporučeni redoslijed rada

1. **Inventory**: `.\Invoke-SoaConversion.ps1 -Action Inventory -Category Group -GroupType Security`
2. **Pregled CSV-a**: provjerite stupce `CurrentSoaState`, `EligibleForConvert`, `Note`, `SoaNotes` (za grupe).
   Obrišite redove koje ne želite konvertirati (vidi *CSV i Excel* niže, zbog encodinga).
3. **Convert s `-WhatIf`**: validira CSV, radi svježi backup i pokazuje što bi se promijenilo. Ništa se ne mijenja.
   `.\Invoke-SoaConversion.ps1 -Action Convert -Category Group -InputCsv .\groups.csv -WhatIf`
4. **Convert s `-MaxObjects 1`**: jedan pilot objekt.
5. **Post-check pilota**: automatski Verify u istom runu, zatim ručno:
   - Entra admin center: objekt više nije "Windows Server AD DS" source; Audit log aktivnost
     *Change Source of Authority from AD to cloud*.
   - Nakon sljedećeg Connect Sync ciklusa: Event ID **6956** na Connect serveru za taj objekt (promjena
     iz AD-a je blokirana).
   - Za User: probna prijava pilot korisnika (vidi upozorenja niže).
   - Za ExchangeAttributes: promjena Exchange atributa u EXO i provjera writebacka u AD.
6. **Convert manjeg skupa** (npr. `-MaxObjects 10`), pa ponovno post-check.
7. **Convert ostatka** istog CSV-a. Već konvertirani objekti preskaču se kao `AlreadyCloudManaged`.
8. **Verify** po potrebi kasnije: `-Action Verify -InputCsv <bilo koji CSV s ObjectId>`.
9. **Rollback** samo ako je potrebno (vidi niže).

Za ugniježđene grupe: SOA se ne prenosi na podgrupe, pa se svaka grupa konvertira posebno, počevši od najniže
razine. Stupci `IsNested`, `ParentGroups` i `ChildGroups` u inventoryju pokazuju strukturu.

## Što se događa u svakoj akciji

**Inventory**
- Bez `-InputCsv`: Graph upit `$filter=onPremisesSyncEnabled eq true` s `ConsistencyLevel: eventual` i
  `$count=true`, uz paging preko `@odata.nextLink`. Za mailboxove `Get-Mailbox -Filter 'IsDirSynced -eq $true'`.
  U izvještaj ulaze samo prihvatljivi objekti.
- S `-InputCsv`: jedan red po ObjectId-u; neprihvatljivi i nepronađeni dobiju razlog u stupcu `Note`.
- `isCloudManaged` se čita **po objektu** (`GET .../onPremisesSyncBehavior`), jer se ne može dohvatiti list
  upitom (vidi *Razlike*). Za velike tenante to je jedan dodatni poziv po objektu (kod usera dva, zbog managera).

**Convert / Rollback**
1. CSV mora imati stupce `ObjectId` i `Category`, a `Category` mora biti jednak `-Category` u **svim** redovima.
   U suprotnom skripta staje. Neispravan GUID u bilo kojem redu također zaustavlja skriptu. Duplikati i prazni
   redovi se ignoriraju (uz upozorenje).
2. Svježi inventory služi kao backup (`Backup_Inventory_<Category>.csv`, za grupe i `Backup_GroupMemberships.csv`).
   Ako se backup ne može zapisati, skripta staje prije promjene.
3. Ponovna provjera prihvatljivosti:
   - Convert preskače `AlreadyCloudManaged`, `NotFound`, `NotSynced`, `GroupTypeNotSelected` i greške čitanja.
   - Rollback mijenja samo objekte koji su sada `CloudManaged`. Preskače grupe s cloud članovima
     (`CloudMemberCount > 0`), jer Microsoft traži da se cloud reference uklone prije rollbacka.
4. Upozorenje za kategoriju, sažetak i potvrda upisom `YES` (osjetljivo na velika/mala slova).
   Za **User Convert** traže se dvije potvrde. S `-WhatIf` nema potvrde i ništa se ne mijenja.
5. Promjena objekt po objekt (`ShouldProcess`; s `-Confirm` pita za svaki objekt). Greška na jednom objektu
   ne zaustavlja ostale.
6. Graph 429 i 500/502/503/504: do 6 pokušaja, uz `Retry-After` ako ga servis vrati, inače eksponencijalni
   backoff (najviše 60 s).
7. Automatski Verify: do 5 provjera s pauzom od 20 s (eventual consistency).

**Verify** (samostalno): čita trenutno stanje svakog ObjectId-a iz CSV-a, bez ponavljanja. `StateBefore` se uzima
iz stupca `CurrentSoaState` ili `StateAfter` ulaznog CSV-a, ako postoji.

## Izlazne datoteke

Svaki run kreira `OutputFolder\yyyyMMdd_HHmmss_<Action>_<Category>` (uz sufiks `_1`, `_2`… ako mapa već postoji).
Postojeće datoteke se nikad ne prepisuju.

| Datoteka | Kada |
|---|---|
| `Transcript.log` | uvijek |
| `Inventory_<Category>.csv` | Inventory |
| `GroupMemberships.csv` | Inventory, Group |
| `Backup_Inventory_<Category>.csv` (+ `Backup_GroupMemberships.csv`) | Convert, Rollback |
| `Results_<Action>_<Category>.csv` | Convert, Rollback, Verify |

- Ključ u svim CSV-ovima je `ObjectId` (Entra object ID, za mailboxove `ExternalDirectoryObjectId`).
- Svaki CSV ima `Category` i `CurrentSoaState`. Results CSV ima tražene stupce
  `ObjectId; DisplayName; Category; StateBefore; StateAfter; Status; ErrorMessage; Timestamp` i na kraju
  `CurrentSoaState` (= `StateAfter`), pa se može izravno koristiti kao ulaz za Verify ili Rollback.
- Više vrijednosti u jednoj ćeliji odvojene su s `|`. Složeni objekti (npr. adrese kontakta) zapisani su kao
  kompaktni JSON, a telefoni kontakta kao `tip:broj`.

`CurrentSoaState` vrijednosti:

| Vrijednost | Značenje |
|---|---|
| `OnPremManaged` | Graph: `onPremisesSyncEnabled=true`, `isCloudManaged=false`. EXO: `IsDirSynced=true`, `IsExchangeCloudManaged=false`. Prihvatljivo za Convert. |
| `CloudManaged` | `isCloudManaged=true` / `IsExchangeCloudManaged=true`. |
| `NotSynced` | Graph: `isCloudManaged=false` i `onPremisesSyncEnabled` nije `true` (cloud-only objekt **ili** rollback koji čeka sljedeći sync ciklus). EXO: `IsDirSynced=false`. |
| `NotFound` | Objekt s tim ObjectId-om ne postoji. |
| `Unknown` | Greška pri čitanju (razlog je u `Note` / `ErrorMessage`). |

Backup stupci po kategoriji:
- **ExchangeAttributes**: svi traženi atributi, `UserCertificate`/`UserSMimeCertificate` samo kao
  `HasUserCertificate`/`HasUserSMimeCertificate` (True/False), te `Guid` i `ExchangeGuid`. Dodani su i ostali
  atributi koje Microsoftova tablica navodi kao uređive u EXO nakon prebacivanja (ArchiveName, Audit*,
  SingleItemRecoveryEnabled, LitigationHoldDate/Owner, retention hold datumi, RetentionComment/Url,
  MailTipTranslations, moderacija, Resource*, ImmutableId) te `AssistantName` iz `Get-User`. **Nisu u backupu:**
  mailbox permissions (`Get-MailboxPermission`) i junk e-mail postavke; izvezite ih zasebno ako ih mijenjate u
  cloudu. On-prem DN, SAM i
  ImmutableId nisu u EXO, a skripta za ovu kategoriju ne spaja Graph. Objekt u AD-u pronađite po UPN-u ili
  PrimarySmtpAddress.
- **User**: svi traženi atributi, `managerUserPrincipalName`, `extensionAttribute1-15`, svi `onPremises*`.
- **Group**: traženi Graph atributi, `extensionAttribute1-15`, broj članova, cloud članova, podgrupa i nadgrupa,
  vlasnici, `SoaNotes`, te `Exo_*` stupci za mail-enabled grupe. Graph v1.0 za grupe nema
  `onPremisesDistinguishedName` ni `onPremisesImmutableId`; za pronalazak u AD-u koristite
  `onPremisesSamAccountName` / `onPremisesSecurityIdentifier`.
- **Contact**: traženi atributi + `Exo_ExternalEmailAddress`, `Exo_HiddenFromAddressListsEnabled`,
  `Exo_EmailAddresses`. Graph za kontakte nema on-prem DN.

## CSV i Excel

CSV-ovi se zapisuju kao `-Encoding Unicode` (UTF-16 LE s BOM-om), kako ste tražili, a import čita isti encoding.
Skripta odbija ulazni CSV bez UTF-16 BOM-a, umjesto da tiho pročita krive znakove.

**Moguć problem (nije testiran u Excelu):** Excel UTF-16 datoteke otvara kao *Unicode Text*, što za Excel znači
**tab** kao separator. Dijakritici (č, ć, ž, š, đ) bit će ispravni, ali na dvoklik se svi stupci mogu pojaviti u
jednom stupcu. Preporuke:
- Za pregled: Excel → *Data → From Text/CSV*, File Origin *Unicode (UTF-16 LE)*, Delimiter *Semicolon*.
- Za brisanje redova prije Converta najsigurnije je koristiti tekst editor (VS Code, Notepad++) i spremiti kao
  UTF-16 LE. Ako se datoteka sprema iz Excela kao *Unicode Text*, ona postaje tab-delimited; tada pokrenite
  skriptu s ``-Delimiter "`t"``. Skripta jasno javi ako ne pronađe stupac `ObjectId`.
- Ako vam UTF-8 s BOM-om više odgovara (Excel s hrvatskim regionalnim postavkama ga na dvoklik otvara ispravno
  i s `;`), promijenite jednu varijablu na vrhu skripte: `$CsvEncoding = 'utf8BOM'`. Import tada mora koristiti
  isti encoding, a provjera BOM-a u `Read-InputCsv` treba se prilagoditi.

## Upozorenja koja skripta prikazuje

**ExchangeAttributes Convert** (prema Microsoftu):
- Od zadnje on-prem promjene (`Set-RemoteMailbox`) ili premještanja mailboxa mora proći puni sync ciklus + 24 h
  (npr. sync 6 h = ukupno 30 h).
- Nakon premještanja mailboxa: prvo potvrdite da je premještanje gotovo i da je on-prem objekt RemoteMailbox.
  Zatim pričekajte da se atributi nakon premještanja sinkroniziraju u Entra ID, pa tek onda počnite brojati 24 h.
  Prerana konverzija može dovesti do toga da writeback zastarjelim cloud recipient-type vrijednostima prepiše
  ispravne remote-mailbox vrijednosti u AD-u.
- Skripta to ne može provjeriti.
- Writeback u AD (Cloud Sync) postoji samo za CustomAttribute1-15, ExtensionCustomAttribute1-5, mail
  (WindowsEmailAddress), proxyAddresses i recipient type. Ostalo (npr. HiddenFromAddressListsEnabled,
  forwarding, moderacija) ostaje samo u cloudu, a nakon rollbacka se vraća na vrijednost iz AD-a.
- Prije vraćanja mailboxa on-prem (offboarding) treba prvo napraviti Rollback (`IsExchangeCloudManaged = $false`),
  inače sinkronizacija pri offboardingu puca.

**User Convert**:
- Korisnik postaje cloud authoritative i promjene iz AD-a se više ne primjenjuju (Connect Sync bilježi event 6956).
- Konverzija briše `IsExchangeCloudManaged` stanje.
- Prijava i lozinka, prema Microsoftovoj dokumentaciji:
  - korisnici federirani preko **AD FS nisu podržani**,
  - korisnici ne smiju ovisiti o on-prem autentifikaciji lozinkom (LDAP bind, Kerberos s lozinkom),
  - preporučena je passwordless autentifikacija (Windows Hello for Business s **Cloud Kerberos Trust**, FIDO2),
  - kod third-party federacije AD račun i lozinku treba održavati ručno.
- Svi mailboxovi moraju biti u EXO prije User SOA.
- `onPremises*` atributi ostaju na cloud objektu i treba ih ručno održavati ako je potreban pristup on-prem
  resursima.

**Group Convert**: SOA se ne prenosi na ugniježđene grupe. DL i MESG se nakon konverzije uređuju samo u EXO i
nemaju writeback u AD. `extensionAttribute1-15` nisu podržani na konvertiranim grupama.

**Contact Convert**: potreban je Connect Sync 2.5.79.0+.

**Rollback (sve kategorije)**: sljedeći sync prepisuje cloud vrijednosti vrijednostima iz AD-a. Ono što treba
sačuvati nalazi se u backup CSV-u (putanja se ispisuje). Skripta ne vraća atribute automatski. Rollback je gotov
tek nakon sljedećeg Connect Sync ciklusa.

**User Rollback, dodatno**: prema Microsoftu, prije user rollbacka treba isključiti tenant feature
`blockCloudObjectTakeoverThroughHardMatchEnabled`
(`PATCH /beta/directory/onPremisesSynchronization/{id}`), a nakon što Connect Sync preuzme objekt treba ga
ponovno uključiti. To je tenant-wide postavka, pa je skripta **ne mijenja**, nego samo upozorava. Također treba
ukloniti cloud reference (članstva u konvertiranim grupama, access packages).

## Razlike između zahtjeva i Microsoftove dokumentacije

Provjereno na learn.microsoft.com 29. 10. 2026:
- *Cloud-based management of Exchange attributes for Remote Mailboxes in hybrid environments* (ažurirano
  2026-09-28),
- *Configure user SOA* (2026-04-15), *Configure Group SOA* (2025-11-19), *Transfer user SOA* (2026-09-24),
- Graph *Get/Update onPremisesSyncBehavior* i *onPremisesSyncBehavior resource type*,
- *Decommission the last Exchange Server after transferring SOA to cloud*.

Microsoftov blog na techcommunity.microsoft.com nije bio dostupan, ali službena Exchange stranica pokriva sve
što je bilo potrebno.

1. **v1.0 ili beta.** Entra how-to članci (User/Group/Contact SOA) koriste `https://graph.microsoft.com/v1.0/...`.
   Graph API referenca dokumentira `onPremisesSyncBehavior` (get/update/resource) **samo pod beta**; i kad se na
   learn.microsoft.com odabere v1.0, stranica prikazuje beta sadržaj.
   Skripta koristi v1.0, a ako v1.0 vrati *"Resource not found for the segment"*, automatski prelazi na beta za
   SOA endpoint. U logu se ispisuje koja je verzija korištena.
2. **Stanje nakon konverzije.** `onPremisesSyncEnabled` nakon konverzije **ne** ostaje `true`, nego postaje
   **`null`**. Microsoftova tablica:

   | Korak | isCloudManaged | onPremisesSyncEnabled |
   |---|---|---|
   | Objekt sinkroniziran iz AD-a | false | true |
   | SOA prebačen u cloud | true | null |
   | Rollback | false | null (dok sync klijent ne preuzme objekt) |
   | Cloud-native objekt | false | null |

   Posljedice: Inventory bez CSV-a ne vidi konvertirane objekte, Verify se oslanja isključivo na `isCloudManaged`,
   a rollback i cloud-native objekt izgledaju isto dok sync ne preuzme objekt (stanje `NotSynced`).
3. **`$select` i `$expand`.** GET `onPremisesSyncBehavior` prema API referenci **ne podržava query parametre**
   (Entra članci ipak koriste `?$select=isCloudManaged`). Skripta zato ne šalje `$select`; odgovor ionako sadrži
   samo `id` i `isCloudManaged`. `$expand` na list upitima nije dokumentiran, a advanced queries ionako ne
   podržavaju `$expand`. Beta dokumentira samo `$filter=onPremisesSyncBehavior/isCloudManaged eq true` uz
   advanced query, a to ne daje vrijednost po objektu. Zato se `isCloudManaged` čita po objektu.
4. **Lozinka i prijava nakon User SOA.** Dokumentacija ne opisuje izričito što se događa s password hash syncom.
   Navodi preduvjete: bez AD FS federacije, bez aplikacija s autentifikacijom lozinkom, passwordless s Cloud
   Kerberos Trust. Upozorenje u skripti prenosi upravo to i preporučuje test s pilot korisnikom.
5. **User rollback** traži isključivanje `blockCloudObjectTakeoverThroughHardMatchEnabled` (tenant-wide), što
   zahtjev ne spominje. Skripta tu postavku ne dira.
6. **Grupe:** prije rollbacka treba ukloniti cloud članove. Skripta takve grupe preskače. Na konvertiranim
   grupama nisu podržani `extensionAttribute1-15`. DL i MESG se nakon konverzije ne mogu uređivati kroz Graph ni
   Entra admin center, nego samo u EXO.
7. **Verzije Connect Synca:** minimalno 2.5.76.0 (User/Group) i 2.5.79.0 (Contact).
8. **ExchangeAttributes: zahtjev je potvrđen.** Dokumentacija potvrđuje:
   - pravilo sync ciklus + 24 h (uz dodatna pravila za premještanje mailboxa, vidi *Upozorenja*),
   - da se `-IsExchangeCloudManaged` *"should not be used together with any other parameters"*,
   - da certificate-based app-only nije podržan (*"an app-only session doesn't provide the delegated user
     authorization required to update the IsExchangeCloudManaged property in Microsoft Entra ID"*),
   - da User SOA *"clears the user's IsExchangeCloudManaged state"*.

   Dodatno iz dokumentacije:
   - zastavica vrijedi samo za mailboxove korisnika s `IsDirSynced = true`; za mail-enabled grupe i mail
     kontakte koristi se Group/Contact SOA,
   - nakon User SOA `IsDirSynced` postaje `False` (skripta takav mailbox prikazuje kao `NotSynced`),
   - writeback u AD radi samo za dio atributa (vidi *Upozorenja*).

   Microsoftovi primjeri koriste `-Identity <User>` / `Alias`. Dokumentacija ne navodi izričito da `-Identity`
   prihvaća `ExternalDirectoryObjectId`. Skripta ga koristi kako je traženo, a automatski Verify potvrđuje
   promjenu. Ako EXO ne razriješi ObjectId, objekt završava kao `Failed` i ništa se ne mijenja; to se vidi već
   na pilotu s `-MaxObjects 1`.

## Testiranje

Skripta je testirana samo protiv lažnih (stub) `Microsoft.Graph.Authentication` i `ExchangeOnlineManagement`
modula. Testirano je: Inventory (sve kategorije), Convert, Rollback, Verify, `-WhatIf`, `-MaxObjects`, neslaganje
kategorije, CSV koji nije UTF-16, otkazivanje, retry na 429, prelazak v1.0 → beta, greška na jednom objektu
dok ostali nastavljaju, i neuspjeli verify. **Nije testirana protiv pravog tenanta** i PSScriptAnalyzer nije
pokrenut. Prije produkcije pokrenite PSScriptAnalyzer i slijedite redoslijed iznad, s jednim pilot objektom.
