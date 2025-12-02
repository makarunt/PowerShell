# Usage Examples

This document provides practical examples and scenarios for using the Exchange Connectors and Certificates Information Script.

## Table of Contents
- [Quick Start](#quick-start)
- [Common Scenarios](#common-scenarios)
- [Advanced Usage](#advanced-usage)
- [Scheduled Tasks](#scheduled-tasks)
- [Output Examples](#output-examples)

## Quick Start

### Run with Default Settings
The simplest way to run the script - gathers info from all servers and exports both CSV and HTML:

```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1
```

**Expected Output:**
```
[2025-12-02 10:30:15] [Info] ======================================
[2025-12-02 10:30:15] [Info] Exchange Connectors and Certificates Information Gathering
[2025-12-02 10:30:15] [Info] ======================================
[2025-12-02 10:30:15] [Info] Gathering information from all Exchange servers...
[2025-12-02 10:30:16] [Info] Found 3 Exchange server(s) to process
[2025-12-02 10:30:16] [Info] Collecting Receive Connector information...
[2025-12-02 10:30:18] [Success]   Processed 12 receive connectors from EXCH01
...
```

## Common Scenarios

### Scenario 1: Quick Health Check of Specific Server

When you need to quickly check a specific Exchange server's configuration:

```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -Server "EXCH01" -ExportHTML
```

**Use Case:**
- Troubleshooting mail flow issues on a specific server
- Quick visual report for management
- Pre-maintenance verification

**What You Get:**
- HTML report opens automatically in browser
- Certificate expiration warnings highlighted
- Connector status at a glance

---

### Scenario 2: Certificate Audit for Compliance

Generate detailed certificate reports for compliance documentation:

```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportCSV -OutputPath "C:\Compliance\Exchange\Certificates"
```

**Use Case:**
- Annual security audits
- Compliance reporting
- Certificate inventory management

**What You Get:**
- Detailed CSV with all certificate properties
- Expiration dates for planning renewals
- Self-signed certificate identification

---

### Scenario 3: Complete Environment Documentation

Create comprehensive documentation of entire Exchange environment:

```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportAll -OutputPath "C:\Documentation\Exchange\$(Get-Date -Format 'yyyy-MM')"
```

**Use Case:**
- Disaster recovery documentation
- Handover documentation for new staff
- Architecture documentation

**What You Get:**
- Complete CSV exports for all components
- Professional HTML report
- Organized in monthly folders

---

### Scenario 4: Pre-Migration Assessment

Before migrating to a new Exchange server or upgrading:

```powershell
# Document current state
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportAll -OutputPath "C:\Migration\Pre-Migration"

# After migration, document new state
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportAll -OutputPath "C:\Migration\Post-Migration"
```

**Use Case:**
- Exchange server migrations
- Version upgrades
- Configuration comparisons

---

### Scenario 5: Multi-Server Comparison

Compare configurations across multiple servers:

```powershell
# Get info from first server
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -Server "EXCH01" -ExportCSV -OutputPath "C:\Reports\EXCH01"

# Get info from second server
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -Server "EXCH02" -ExportCSV -OutputPath "C:\Reports\EXCH02"

# Compare CSV files to identify differences
```

**Use Case:**
- Configuration standardization
- Troubleshooting inconsistencies
- Load balancing verification

---

### Scenario 6: Certificate Expiration Monitoring

Regular monitoring to prevent certificate expiration issues:

```powershell
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportHTML -OutputPath "C:\Reports\Certificates"
```

**Use Case:**
- Monthly certificate reviews
- Proactive renewal planning
- Avoiding service disruptions

**Script Output Highlights:**
```
[2025-12-02 10:35:42] [Success]   Certificates: 15
[2025-12-02 10:35:42] [Error]   WARNING: 1 expired certificate(s) found!
[2025-12-02 10:35:42] [Warning]   WARNING: 2 certificate(s) expiring within 30 days!
```

---

## Advanced Usage

### Scheduled Daily Report

Create a scheduled task to run daily and email reports:

```powershell
# Script content for scheduled task
$reportPath = "C:\Reports\Exchange\Daily"
$timestamp = Get-Date -Format "yyyyMMdd"

# Run the collection script
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportHTML -OutputPath $reportPath

# Email the report (customize for your environment)
$htmlReport = Get-ChildItem $reportPath -Filter "*.html" | Sort-Object LastWriteTime -Descending | Select-Object -First 1

Send-MailMessage `
    -To "exchange-admins@company.com" `
    -From "exchange-reports@company.com" `
    -Subject "Daily Exchange Connectors Report - $timestamp" `
    -Body "Attached is today's Exchange connectors and certificates report." `
    -Attachments $htmlReport.FullName `
    -SmtpServer "smtp.company.com"
```

---

### Filter and Analyze Specific Connectors

Using PowerShell to analyze the exported data:

```powershell
# First, run the script
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportCSV

# Then analyze the data
$timestamp = Get-Date -Format "yyyyMMdd"
$receiveConnectors = Import-Csv "ReceiveConnectors_$timestamp*.csv"

# Find disabled connectors
$disabledConnectors = $receiveConnectors | Where-Object {$_.Enabled -eq $false}
Write-Host "Disabled Receive Connectors: $($disabledConnectors.Count)"
$disabledConnectors | Format-Table Server, Name, Identity

# Find connectors without TLS
$noTLS = $receiveConnectors | Where-Object {$_.RequireTLS -eq $false}
Write-Host "`nConnectors without TLS requirement: $($noTLS.Count)"
$noTLS | Format-Table Server, Name, RequireTLS
```

---

### Certificate Expiration Alert Script

Automated alerting for certificate expiration:

```powershell
# Run the collection
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportCSV -OutputPath "C:\Temp"

# Import and analyze certificates
$csvFile = Get-ChildItem "C:\Temp\Certificates_*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$certificates = Import-Csv $csvFile.FullName

# Check for problems
$expired = $certificates | Where-Object {$_.ExpiryStatus -eq "Expired"}
$expiring = $certificates | Where-Object {$_.ExpiryStatus -eq "Expiring Soon"}

if ($expired.Count -gt 0 -or $expiring.Count -gt 0) {
    $message = @"
ALERT: Exchange Certificate Expiration Warning

Expired Certificates: $($expired.Count)
$($expired | Select-Object Server, Subject, NotAfter | Format-Table | Out-String)

Expiring Soon (within 30 days): $($expiring.Count)
$($expiring | Select-Object Server, Subject, NotAfter, DaysUntilExpiry | Format-Table | Out-String)

Action Required: Renew these certificates immediately.
"@

    Send-MailMessage `
        -To "exchange-admins@company.com" `
        -From "exchange-alerts@company.com" `
        -Subject "URGENT: Exchange Certificate Expiration Alert" `
        -Body $message `
        -SmtpServer "smtp.company.com" `
        -Priority High
}
```

---

### Connector Configuration Backup

Use as part of backup procedures:

```powershell
# Create backup directory with date
$backupPath = "C:\Backups\Exchange\Config\$(Get-Date -Format 'yyyyMMdd_HHmmss')"
New-Item -ItemType Directory -Path $backupPath -Force

# Run the script
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportAll -OutputPath $backupPath

# Compress for archival
Compress-Archive -Path "$backupPath\*" -DestinationPath "$backupPath.zip"

Write-Host "Backup completed: $backupPath.zip"
```

---

## Scheduled Tasks

### Example: Weekly Certificate Report

Create a scheduled task using PowerShell:

```powershell
$action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
    -Argument "-ExecutionPolicy Bypass -File C:\Scripts\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportHTML -OutputPath C:\Reports\Weekly"

$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday -At 6am

$principal = New-ScheduledTaskPrincipal -UserId "DOMAIN\ExchangeAdmin" `
    -LogonType Password -RunLevel Highest

Register-ScheduledTask -TaskName "Exchange Certificate Weekly Report" `
    -Action $action `
    -Trigger $trigger `
    -Principal $principal `
    -Description "Generates weekly Exchange connectors and certificates report"
```

---

## Output Examples

### Console Output Example

```
[2025-12-02 10:30:15] [Info] ======================================
[2025-12-02 10:30:15] [Info] Exchange Connectors and Certificates Information Gathering
[2025-12-02 10:30:15] [Info] ======================================
[2025-12-02 10:30:15] [Info] Gathering information from all Exchange servers...
[2025-12-02 10:30:16] [Info] Found 3 Exchange server(s) to process
[2025-12-02 10:30:16] [Info] Output path: C:\Reports
[2025-12-02 10:30:16] [Info] Collecting Receive Connector information...
[2025-12-02 10:30:18] [Success]   Processed 4 receive connectors from EXCH01
[2025-12-02 10:30:20] [Success]   Processed 4 receive connectors from EXCH02
[2025-12-02 10:30:22] [Success]   Processed 4 receive connectors from EXCH03
[2025-12-02 10:30:22] [Info] Collecting Send Connector information...
[2025-12-02 10:30:24] [Success]   Processed 5 send connectors
[2025-12-02 10:30:24] [Info] Collecting Exchange Certificate information...
[2025-12-02 10:30:26] [Success]   Processed 5 certificates from EXCH01
[2025-12-02 10:30:28] [Success]   Processed 5 certificates from EXCH02
[2025-12-02 10:30:30] [Success]   Processed 5 certificates from EXCH03
[2025-12-02 10:30:30] [Info] ======================================
[2025-12-02 10:30:30] [Info] Data Collection Summary:
[2025-12-02 10:30:30] [Success]   Receive Connectors: 12
[2025-12-02 10:30:30] [Success]   Send Connectors: 5
[2025-12-02 10:30:30] [Success]   Certificates: 15
[2025-12-02 10:30:30] [Warning]   WARNING: 2 certificate(s) expiring within 30 days!
[2025-12-02 10:30:30] [Info] ======================================
[2025-12-02 10:30:30] [Info] Exporting data to CSV files...
[2025-12-02 10:30:31] [Success]   Receive Connectors exported to: C:\Reports\ReceiveConnectors_20251202_103031.csv
[2025-12-02 10:30:31] [Success]   Send Connectors exported to: C:\Reports\SendConnectors_20251202_103031.csv
[2025-12-02 10:30:31] [Success]   Certificates exported to: C:\Reports\Certificates_20251202_103031.csv
[2025-12-02 10:30:31] [Info] Generating HTML report...
[2025-12-02 10:30:32] [Success]   HTML report generated: C:\Reports\ExchangeConnectorsAndCertificates_20251202_103032.html
[2025-12-02 10:30:32] [Info] ======================================
[2025-12-02 10:30:32] [Success] Script completed successfully!
[2025-12-02 10:30:32] [Info] ======================================
```

### CSV Output Sample (Certificates)

```csv
Server,Thumbprint,Subject,Issuer,FriendlyName,DomainNames,Services,Status,IsSelfSigned,NotBefore,NotAfter,DaysUntilExpiry,ExpiryStatus
EXCH01,1A2B3C4D...,CN=mail.contoso.com,CN=DigiCert,Contoso Mail,mail.contoso.com; autodiscover.contoso.com,IIS; SMTP,Valid,False,2024-01-15,2025-01-15,44,Valid
EXCH01,5E6F7G8H...,CN=EXCH01,CN=EXCH01,Self-Signed,EXCH01.contoso.local,IMAP; POP,Valid,True,2023-06-01,2028-06-01,926,Valid
```

---

## Tips and Best Practices

### Tip 1: Use Variables for Repeated Tasks
```powershell
$scriptPath = "C:\Scripts\Get-ExchangeConnectorAndCertificateInfo.ps1"
$outputPath = "C:\Reports\Exchange"

# Easy to run with consistent settings
& $scriptPath -ExportAll -OutputPath $outputPath
```

### Tip 2: Combine with Git for Version Control
```powershell
$reportPath = "C:\Reports\Exchange\Configs"
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportCSV -OutputPath $reportPath

# Commit to git for change tracking
cd $reportPath
git add .
git commit -m "Exchange config snapshot $(Get-Date -Format 'yyyy-MM-dd')"
```

### Tip 3: Filter HTML Report in Browser
The HTML report generates searchable content. Use Ctrl+F in your browser to quickly find:
- Specific server names
- Connector names
- Certificate subjects
- Expiring certificates

### Tip 4: Compare Historical Data
```powershell
# Export current state
$today = Get-Date -Format "yyyyMMdd"
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportCSV -OutputPath "C:\Reports\$today"

# Compare with last month
$lastMonth = (Get-Date).AddMonths(-1).ToString("yyyyMMdd")
Compare-Object `
    (Import-Csv "C:\Reports\$lastMonth\SendConnectors_*.csv") `
    (Import-Csv "C:\Reports\$today\SendConnectors_*.csv") `
    -Property Name, Enabled, SmartHosts
```

---

## Troubleshooting Examples

### Example: Script Not Running

```powershell
# Check execution policy
Get-ExecutionPolicy

# If restricted, set to RemoteSigned
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser

# Run script
.\Get-ExchangeConnectorAndCertificateInfo.ps1
```

### Example: Verify Exchange Management Shell

```powershell
# Check if Exchange cmdlets are available
Get-Command Get-ExchangeServer

# If not available, load Exchange snap-in
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
```

### Example: Test Server Connectivity

```powershell
# Before running the script, verify server is accessible
$serverName = "EXCH01"
Test-Connection $serverName -Count 2
Get-ExchangeServer $serverName
```

---

## Integration Examples

### Example: ServiceNow Integration

```powershell
# Run script and post to ServiceNow
.\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportCSV -OutputPath "C:\Temp"

$certs = Import-Csv "C:\Temp\Certificates_*.csv"
$expiring = $certs | Where-Object {$_.DaysUntilExpiry -lt 30 -and $_.DaysUntilExpiry -gt 0}

foreach ($cert in $expiring) {
    # Create ServiceNow incident (customize for your instance)
    $body = @{
        short_description = "Exchange Certificate Expiring: $($cert.Subject)"
        description = "Certificate on $($cert.Server) expires in $($cert.DaysUntilExpiry) days"
        urgency = if($cert.DaysUntilExpiry -lt 15) {1} else {2}
    } | ConvertTo-Json

    Invoke-RestMethod -Uri "https://instance.service-now.com/api/now/table/incident" `
        -Method Post -Body $body -ContentType "application/json" `
        -Headers @{Authorization = "Basic $encodedCredentials"}
}
```

---

For more information, see README.md
