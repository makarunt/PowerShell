# Exchange IMAP/POP3 IP Tracking - Vodič za Pravilno Korištenje

## 🔍 Problem: Gdje se nalazi PRAVA klijentska IP adresa?

U Exchange okruženju s više servera, tok konekcije izgleda ovako:

```
┌──────────┐      ┌─────────────┐      ┌─────────────────┐
│ Klijent  │─────▶│ CAS Server  │─────▶│ Mailbox Server  │
│          │      │ (Frontend)  │      │ (Backend)       │
└──────────┘      └─────────────┘      └─────────────────┘
  Prava IP         Proxy/LB IP          Vidi CAS IP
```

### U IMAP/POP3 logovima na Mailbox serveru:
- **`sIp`** = Mailbox server IP (gdje se nalazi log)
- **`cIp`** = CAS server IP ili Load Balancer IP (**NE prava klijentska IP!**)

### Prava klijentska IP adresa se nalazi u:
1. **IIS logovima na CAS serveru** ← **Preporučeno**
2. HttpProxy logovima na CAS serveru
3. Load Balancer logovima (F5, HAProxy, itd.)

---

## 📋 Dostupne Skripte

### 1️⃣ **Diagnose-ExchangeMailLogs.ps1** - Dijagnostička skripta
**Svrha:** Prikazuje SVA polja iz IMAP/POP3 logova i traži IP adrese

**Koristi kada:** Želite vidjeti što se točno nalazi u logovima

```powershell
.\Diagnose-ExchangeMailLogs.ps1
```

**Output:**
- Detaljni ispis svih polja iz autentikacijskih zapisa
- Traži IP adrese u `context` polju
- Kreira CSV za daljnju analizu

---

### 2️⃣ **Parse-IIS-ClientIPs.ps1** - IIS Log Parser ⭐ **PREPORUČENO**
**Svrha:** Parsira IIS logove na CAS serveru i izvlači PRAVE klijentske IP adrese

**Koristi kada:** Trebate prave klijentske IP adrese (ne CAS server IP)

```powershell
# Ako ste NA CAS serveru:
.\Parse-IIS-ClientIPs.ps1

# Ako ste NA Mailbox serveru, ali možete pristupiti CAS serveru:
# 1. Kopirajte skriptu na CAS server
# 2. Pokrenite je tamo, ILI
# 3. Modificirajte $IISLogPath u skripti da pokazuje na network share:
#    $IISLogPath = "\\CAS-SERVER\c$\inetpub\logs\LogFiles\W3SVC1"
```

**Output:**
- CSV s pravim klijentskim IP adresama
- Username, ClientIP, LastSeen, LastURI, LastStatus

---

### 3️⃣ **Parse-ExchangeMailLogs-Fixed.ps1** - IMAP/POP3 Parser
**Svrha:** Parsira IMAP/POP3 logove (ali daje CAS server IP, ne klijentsku IP)

**Koristi kada:** Trebate vidjeti koje CAS servere koriste korisnici

```powershell
.\Parse-ExchangeMailLogs-Fixed.ps1
```

**NAPOMENA:** Ova skripta će dati IP adrese CAS servera, ne pravih klijenata!

---

## 🎯 Preporučeni Pristup

### Scenarij 1: Imate pristup CAS serveru

```powershell
# Pokrenite na CAS serveru:
.\Parse-IIS-ClientIPs.ps1
```

✅ **Rezultat:** Prave klijentske IP adrese

---

### Scenarij 2: Nemate direktan pristup CAS serveru

#### Opcija A: Remote pristup IIS logovima
```powershell
# Modificirajte Parse-IIS-ClientIPs.ps1:
$IISLogPath = "\\YOUR-CAS-SERVER\c$\inetpub\logs\LogFiles\W3SVC1"
.\Parse-IIS-ClientIPs.ps1
```

#### Opcija B: Kopirajte IIS logove na lokalni server
```powershell
# 1. Kopiraj logove
Copy-Item "\\CAS-SERVER\c$\inetpub\logs\LogFiles\W3SVC1\*.log" -Destination "C:\Temp\IISLogs\"

# 2. Modificiraj skriptu
$IISLogPath = "C:\Temp\IISLogs"

# 3. Pokreni
.\Parse-IIS-ClientIPs.ps1
```

#### Opcija C: HttpProxy logovi (alternativa)
```powershell
# HttpProxy logovi na CAS serveru također mogu sadržavati klijentske IP
$HttpProxyPath = "C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\Rps"
# Slično kao IIS parser, ali za HttpProxy logove
```

---

## 📁 Lokacije Logova

### Na Mailbox Serveru:
```
IMAP: C:\Program Files\Microsoft\Exchange Server\V15\Logging\Imap4\
POP3: C:\Program Files\Microsoft\Exchange Server\V15\Logging\Pop3\
```
⚠️ Sadrže CAS server IP, ne klijentsku IP

### Na CAS Serveru (Frontend):
```
IIS:        C:\inetpub\logs\LogFiles\W3SVC*\
HttpProxy:  C:\Program Files\Microsoft\Exchange Server\V15\Logging\HttpProxy\
```
✅ Sadrže PRAVE klijentske IP adrese

---

## 🔧 Troubleshooting

### Problem: "Nema log datoteka"
**Rješenje:** Provjerite:
1. Da li skripta radi na pravom serveru (CAS za IIS logove)?
2. Da li su putanje točne?
3. Da li korisnik ima permissions za čitanje logova?

### Problem: "Vidim samo interne IP adrese (10.x.x.x)"
**Rješenje:** To su vjerojatno:
- CAS server IP adrese (ako gledate IMAP/POP3 logove na Mailbox serveru)
- Load balancer backend IP (ako gledate IIS logove iza load balancera)

**Provjera:** Potvrdite da li su te IP adrese vaši serveri:
```powershell
Get-ExchangeServer | Select-Object Name, ServerRole, AdminDisplayVersion
Get-NetIPAddress | Where-Object { $_.IPAddress -like "10.150.16.*" }
```

### Problem: Vidim Load Balancer IP umjesto klijentske IP
**Rješenje:** Provjerite `X-Forwarded-For` header u IIS logovima.
Možda trebate:
1. Enableati X-Forwarded-For logging u IIS-u
2. Parsirati taj header u logovima
3. Ili pristupiti load balancer logovima direktno

---

## 📊 Kako Interpretirati Rezultate

### Ako koristite **Parse-IIS-ClientIPs.ps1**:
```csv
UserName,ClientIP,LastSeen,LastURI,LastStatus
john.doe,203.0.113.45,2025-11-03 14:32:01,/Microsoft-Server-ActiveSync,200
jane.smith,198.51.100.88,2025-11-03 14:35:12,/ews/exchange.asmx,200
```
✅ **203.0.113.45** i **198.51.100.88** su PRAVE klijentske IP adrese

### Ako koristite **Parse-ExchangeMailLogs-Fixed.ps1**:
```csv
UserName,IPAddress,LastSeen
john.doe,10.150.16.125,2025-11-03 14:32:01
jane.smith,10.150.16.126,2025-11-03 14:35:12
```
⚠️ **10.150.16.x** su vjerojatno CAS server IP adrese, ne klijentske

---

## 🆘 Dodatna Pomoć

### Identifikacija servera po IP adresi:
```powershell
# Provjeri što je na određenoj IP
Test-Connection -ComputerName 10.150.16.125 -Count 1
Resolve-DnsName 10.150.16.125

# Provjeri sve Exchange servere
Get-ExchangeServer | Select-Object Name, ServerRole, AdminDisplayVersion, InternetFqdn
```

### Provjera Load Balancera:
```powershell
# Ako koristite F5 ili HAProxy, provjerite:
# - Virtual Server IP (ovo vide klijenti)
# - Backend Pool IPs (ovo vide CAS serveri)
```

---

## 📝 Zaključak

| Što trebate? | Koja skripta? | Gdje pokrenuti? |
|--------------|---------------|-----------------|
| **Prave klijentske IP** | `Parse-IIS-ClientIPs.ps1` | Na CAS serveru |
| **Dijagnostika logova** | `Diagnose-ExchangeMailLogs.ps1` | Bilo gdje |
| **CAS server analiza** | `Parse-ExchangeMailLogs-Fixed.ps1` | Na Mailbox serveru |

---

**Autor:** Claude Code
**Verzija:** 1.0
**Datum:** 2025-12-04
