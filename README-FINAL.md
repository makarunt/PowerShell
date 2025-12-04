# 🎯 Kako Dobiti PRAVE Klijentske IP Adrese iz Exchange IMAP/POP3 Logova

## ❌ Zašto originalna skripta ne radi pravilno?

### Problem #1: Parsira sve komande, ne samo autentikaciju
Originalna skripta je uzimala IP iz **bilo koje komande** gdje se pojavljuje username:
- `select INBOX`
- `uid+search`
- `stat`
- Itd.

**Rješenje:** Nova skripta uzima IP samo iz **autentikacijskih komandi** (`authenticate` za IMAP, `auth` za POP3)

### Problem #2: Čita logove sa Mailbox servera umjesto CAS servera ⚠️ **GLAVNI PROBLEM**

U Exchange okruženju sa više servera:

```
┌──────────────┐         ┌─────────────────┐         ┌──────────────────┐
│   Klijent    │────────▶│   CAS Server    │────────▶│ Mailbox Server   │
│              │         │   (Frontend)    │         │   (Backend)      │
│ 203.0.113.45 │         │  10.150.16.125  │         │  10.150.16.111   │
└──────────────┘         └─────────────────┘         └──────────────────┘
  ↑ PRAVA IP                ↑ CAS Server IP             ↑ MBX Server IP
```

**Na Mailbox Serveru (backend):**
- `sIp` = 10.150.16.111 (Mailbox server IP)
- `cIp` = 10.150.16.125 (CAS server IP) ← **Ovo NIJE klijentska IP!**

**Na CAS Serveru (frontend):**
- `sIp` = 10.150.16.125 (CAS server IP)
- `cIp` = 203.0.113.45 (PRAVA klijentska IP) ← **✅ OVO TREBATE!**

---

## 🔍 Korak 1: Identifikacija - Gdje se nalazite?

Pokrenite dijagnostičku skriptu koja će vam reći da li gledate CAS ili Mailbox logove:

```powershell
.\Identify-ExchangeIPs.ps1
```

### Što skripta radi:
1. Analizira IP adrese u logovima
2. Provjerava jesu li privatne (10.x, 172.16-31.x, 192.168.x) ili javne
3. Pokušava DNS rezoluciju
4. Identificira Exchange servere

### Interpretacija rezultata:

#### Ako vidite SAMO privatne IP (10.x, 172.x, 192.168.x):
```
⚠️ Privatne IP adrese (potencijalni interni serveri):
IPAddress      Occurrences  Hostname
---------      -----------  --------
10.150.16.125  450          cas-server.domena.local
10.150.16.126  380          cas-server2.domena.local

⚠️ Ovo su vjerovatno CAS serveri ili load balanceri!
```

**✅ ZAKLJUČAK:** Nalazite se na **Mailbox serveru** i trebate pristupiti logovima na **CAS serveru**!

#### Ako vidite javne IP adrese (ne 10.x, ne 192.168.x):
```
✅ Javne IP adrese (potencijalni eksterni klijenti):
IPAddress       Occurrences  Hostname
---------       -----------  --------
203.0.113.45    125          client-host.isp.com
198.51.100.88   89           mobile-device.carrier.com

✅ Ovo su vjerovatno PRAVE klijentske IP adrese!
```

**✅ ZAKLJUČAK:** Nalazite se na **CAS serveru** i možete koristiti `Parse-ExchangeMailLogs-Fixed.ps1`!

---

## 📋 Korak 2: Dobiti prave klijentske IP adrese

### Scenarij A: Trenutno ste na Mailbox serveru (backend)

Trebate pristupiti logovima na CAS (Frontend) serveru:

```powershell
# 1. Identificirajte CAS servere
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
Get-ExchangeServer | Where-Object { $_.ServerRole -like "*ClientAccess*" }

# Primjer output:
# Name          ServerRole                     AdminDisplayVersion
# ----          ----------                     -------------------
# CAS-SERVER1   ClientAccess, Mailbox          Version 15.0 (Build 1497.2)

# 2. OPCIJA 1: Kopirajte skriptu na CAS server i pokrenite tamo
Copy-Item .\Parse-ExchangeMailLogs-Fixed.ps1 -Destination "\\CAS-SERVER1\C$\Temp\"
Invoke-Command -ComputerName CAS-SERVER1 -ScriptBlock {
    cd C:\Temp
    .\Parse-ExchangeMailLogs-Fixed.ps1
}

# 3. OPCIJA 2: Modificirajte putanje u skripti da pokazuju na CAS server
# Otvorite Parse-ExchangeMailLogs-Fixed.ps1 i promijenite:

$ImapLogPath = "\\CAS-SERVER1\C$\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4"
$Pop3LogPath = "\\CAS-SERVER1\C$\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3"

# Zatim pokrenite:
.\Parse-ExchangeMailLogs-Fixed.ps1
```

### Scenarij B: Trenutno ste na CAS serveru (frontend)

Možete direktno pokrenuti skriptu:

```powershell
.\Parse-ExchangeMailLogs-Fixed.ps1
```

### Scenarij C: Imate sve role na jednom serveru (Mailbox + CAS)

Provjerite da li Exchange logira klijentske IP direktno:

```powershell
# Pokrenite dijagnostiku:
.\Identify-ExchangeIPs.ps1

# Ako vidite javne IP, možete koristiti:
.\Parse-ExchangeMailLogs-Fixed.ps1
```

---

## 🧪 Korak 3: Verifikacija rezultata

Nakon što pokrenete skriptu, provjerite rezultate:

```powershell
# Otvorite generirani CSV
Import-Csv "C:\Temp\Span\Exchange_IMAP_Unique_Auth_*.csv" | Select-Object -First 10

# Provjerite da li su IP adrese klijentske (javne) ili serverske (privatne)
# DOBRE IP adrese (klijentske):
#   - Javne IP: 203.0.113.x, 198.51.100.x, itd.
#   - Možda privatne iz vaše korporativne mreže ako su klijenti interni

# LOŠE IP adrese (serverske):
#   - IP adrese vaših Exchange servera
#   - IP adrese load balancera
```

### Brza provjera pojedinačne IP:

```powershell
# Provjerite da li je IP vaš server:
Test-Connection -ComputerName 10.150.16.125 -Count 1
Resolve-DnsName 10.150.16.125

# Provjerite sve Exchange servere:
Get-ExchangeServer | ForEach-Object {
    $Name = $_.Name
    $IPs = [System.Net.Dns]::GetHostAddresses($_.Fqdn) | Select-Object -ExpandProperty IPAddressToString
    Write-Host "$Name : $($IPs -join ', ')"
}
```

---

## 📊 Razumijevanje Strukture Logova

### IMAP/POP3 Log Format:

```csv
dateTime,sessionId,seqNumber,sIp,cIp,user,duration,command,parameters,context
```

### Polja:

| Polje | Opis | Primjer |
|-------|------|---------|
| `dateTime` | Vrijeme zapisa | 2025-11-03T01:00:00.008Z |
| `sIp` | **Server IP** (gdje je log) | 10.150.16.111:1993 |
| `cIp` | **Client IP** (tko se spaja) | 10.150.16.125:8706 |
| `user` | Username | designbevgroup |
| `command` | IMAP/POP3 komanda | authenticate, auth, select, stat |

### ⚠️ VAŽNO:

- `cIp` polje sadrži "klijentsku" IP **iz perspektive servera gdje je log**
- Na **Mailbox serveru**: `cIp` = CAS server (jer se CAS spaja na Mailbox)
- Na **CAS serveru**: `cIp` = Pravi klijent (jer se klijent spaja na CAS)

---

## 🛠️ Dostupne Skripte

### 1. `Identify-ExchangeIPs.ps1` - **Pokreni PRVO**
**Svrha:** Identificira jesu li IP u logovima serveri ili klijenti

```powershell
.\Identify-ExchangeIPs.ps1
```

**Output:**
- Lista svih IP iz logova
- Klasifikacija (privatna/javna)
- DNS rezolucija
- Identifikacija Exchange servera
- Preporuke što dalje

---

### 2. `Parse-ExchangeMailLogs-Fixed.ps1` - **Glavni parser**
**Svrha:** Parsira IMAP/POP3 logove i izvlači User/IP parove

**Prije pokretanja:**
1. Pokreni `Identify-ExchangeIPs.ps1` da potvrdiš da si na CAS serveru
2. Ili modificiraj putanje da pokazuju na CAS server logove

```powershell
# Ako si na CAS serveru:
.\Parse-ExchangeMailLogs-Fixed.ps1

# Ako si na Mailbox serveru, modificiraj skripta prvo:
# Promijeni $ImapLogPath i $Pop3LogPath da pokazuju na \\CAS-SERVER\C$\...
```

**Output:**
- `Exchange_IMAP_Unique_Auth_YYYYMMDD.csv`
- `Exchange_POP3_Unique_Auth_YYYYMMDD.csv`
- `Exchange_Combined_Unique_Auth_YYYYMMDD.csv`

---

### 3. `Diagnose-ExchangeMailLogs.ps1` - **Dijagnostika**
**Svrha:** Prikazuje SVA polja iz autentikacijskih zapisa za debugging

```powershell
.\Diagnose-ExchangeMailLogs.ps1
```

**Koristi kada:**
- Nisi siguran što se nalazi u logovima
- Trebaš vidjeti context ili parameters polja
- Debug problem

---

## 🔧 Troubleshooting

### Problem: "Vidim samo 10.x.x.x IP adrese"

**Dijagnoza:**
```powershell
.\Identify-ExchangeIPs.ps1
```

**Rješenje:**
1. Provjerite da li su to vaši Exchange serveri
2. Ako jesu, pristupite logovima na CAS serveru
3. Ili pokrenite skriptu direktno na CAS serveru

---

### Problem: "Nema log datoteka"

**Mogući razlozi:**
1. Pogrešna putanja
2. Nema autentikacija u zadanom periodu
3. IMAP/POP3 nisu omogućeni na tom serveru

**Provjera:**
```powershell
# Provjeri da li postoje logovi
Get-ChildItem "C:\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4" -Filter "*.LOG"

# Provjeri da li su servisi uključeni
Get-Service | Where-Object { $_.Name -like "*IMAP*" -or $_.Name -like "*POP*" }

# Provjeri Exchange servise
Get-Service MSExchangeIMAP4, MSExchangeIMAP4BE, MSExchangePOP3, MSExchangePOP3BE
```

---

### Problem: "Ne mogu pristupiti CAS serveru"

**Rješenja:**

#### Opcija 1: PSRemoting
```powershell
# Omogući PSRemoting na CAS serveru (ako već nije)
Invoke-Command -ComputerName CAS-SERVER -ScriptBlock { Enable-PSRemoting -Force }

# Pokreni skriptu remote
$Session = New-PSSession -ComputerName CAS-SERVER
Invoke-Command -Session $Session -FilePath .\Parse-ExchangeMailLogs-Fixed.ps1
Remove-PSSession $Session
```

#### Opcija 2: Kopiraj logove lokalno
```powershell
# Kopiraj logove sa CAS servera
$Source = "\\CAS-SERVER\C$\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4"
$Dest = "C:\Temp\CAS_Logs\Imap4"
Copy-Item -Path $Source -Destination $Dest -Recurse

# Modificiraj skriptu da čita lokalne logove
# $ImapLogPath = "C:\Temp\CAS_Logs\Imap4"
.\Parse-ExchangeMailLogs-Fixed.ps1
```

---

## ✅ Checklist za Uspješno Dobivanje Pravih IP Adresa

- [ ] Pokreni `Identify-ExchangeIPs.ps1` da identificiraš tip servera
- [ ] Ako si na Mailbox serveru, identificiraj CAS server(e)
- [ ] Pristup logovima na CAS serveru (direktno ili remote)
- [ ] Pokreni `Parse-ExchangeMailLogs-Fixed.ps1` na CAS server logovima
- [ ] Verifikuj da su IP adrese u rezultatu javne (ne 10.x, ne 192.168.x)
- [ ] Ako su i dalje privatne, možda imate load balancer ispred CAS-a

---

## 🆘 Dodatne Napomene

### Load Balancer Ispred CAS Servera

Ako imate arhitekturu:
```
Klijent → Load Balancer → CAS Server → Mailbox Server
```

Čak i na CAS serveru, `cIp` može biti Load Balancer IP, ne pravi klijent!

**Rješenje:**
1. Provjerite da li CAS logira X-Forwarded-For podatke
2. Ili pristupite Load Balancer logovima (F5, HAProxy, itd.)

### Interni Klijenti

Ako su svi vaši klijenti u **istoj internoj mreži** (npr. korporativna VPN 10.x.x.x):
- Onda će `cIp` biti privatne IP (10.x), ali to **jesu** prave klijentske IP
- Provjerite da li te IP adrese odgovaraju klijentskim računalima, ne serverima

---

## 📞 Brzi Vodič (TL;DR)

```powershell
# 1. Identificiraj gdje si:
.\Identify-ExchangeIPs.ps1

# 2. Ako vidiš privatne IP i server kaže "ovo su CAS serveri":
#    → Modificiraj skriptu da pokazuje na CAS server logove

# 3. Ako vidiš javne IP:
#    → Direktno pokreni parser:
.\Parse-ExchangeMailLogs-Fixed.ps1

# 4. Verifikuj rezultate:
Import-Csv "C:\Temp\Span\Exchange_IMAP_Unique_Auth_*.csv" | Select-Object -First 10
```

---

**Autor:** Claude Code
**Verzija:** 2.0 (FINALNA)
**Datum:** 2025-12-04
**Note:** **NEMOJ koristiti IIS logove** - oni ne sadrže IMAP/POP3 podatke!
