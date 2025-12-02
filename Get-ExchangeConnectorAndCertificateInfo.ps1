<#
.SYNOPSIS
    Gathers information about Exchange connectors and certificates.

.DESCRIPTION
    This script collects detailed information about:
    - Receive Connectors
    - Send Connectors
    - Exchange Certificates

    Data can be exported to CSV files and/or HTML report.
    Can target specific Exchange server or all Exchange servers.

.PARAMETER Server
    Specific Exchange server name. If not specified, all Exchange servers will be queried.

.PARAMETER OutputPath
    Path where output files will be saved. Default is current directory.

.PARAMETER ExportCSV
    Export data to CSV files.

.PARAMETER ExportHTML
    Generate HTML report.

.PARAMETER ExportAll
    Export both CSV and HTML formats.

.EXAMPLE
    .\Get-ExchangeConnectorAndCertificateInfo.ps1 -ExportAll
    Gathers information from all Exchange servers and exports to both CSV and HTML.

.EXAMPLE
    .\Get-ExchangeConnectorAndCertificateInfo.ps1 -Server "EX01" -ExportCSV -OutputPath "C:\Reports"
    Gathers information from server EX01 and exports to CSV in C:\Reports.

.EXAMPLE
    .\Get-ExchangeConnectorAndCertificateInfo.ps1 -Server "EX01" -ExportHTML
    Gathers information from server EX01 and generates HTML report.

.NOTES
    Requires Exchange Management Shell to be available.
    Run with appropriate Exchange administrator permissions.
#>

[CmdletBinding(DefaultParameterSetName='ExportType')]
param(
    [Parameter(Mandatory=$false)]
    [string]$Server,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = ".",

    [Parameter(ParameterSetName='ExportType')]
    [switch]$ExportCSV,

    [Parameter(ParameterSetName='ExportType')]
    [switch]$ExportHTML,

    [Parameter(ParameterSetName='ExportAll')]
    [switch]$ExportAll
)

#region Helper Functions

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('Info','Warning','Error','Success')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $color = switch($Level) {
        'Info'    { 'Cyan' }
        'Warning' { 'Yellow' }
        'Error'   { 'Red' }
        'Success' { 'Green' }
    }

    Write-Host "[$timestamp] [$Level] $Message" -ForegroundColor $color
}

function Test-ExchangeManagementShell {
    try {
        $null = Get-Command Get-ExchangeServer -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Get-ExchangeServerList {
    param([string]$ServerName)

    try {
        if ($ServerName) {
            Write-Log "Targeting specific server: $ServerName"
            $servers = @(Get-ExchangeServer -Identity $ServerName -ErrorAction Stop)
        }
        else {
            Write-Log "Gathering information from all Exchange servers..."
            $servers = Get-ExchangeServer | Where-Object {$_.IsMailboxServer -or $_.IsHubTransportServer}
        }
        return $servers
    }
    catch {
        Write-Log "Error getting Exchange server list: $($_.Exception.Message)" -Level Error
        throw
    }
}

#endregion

#region Data Collection Functions

function Get-ReceiveConnectorInfo {
    param([array]$Servers)

    Write-Log "Collecting Receive Connector information..."
    $receiveConnectors = @()

    foreach ($srv in $Servers) {
        try {
            $connectors = Get-ReceiveConnector -Server $srv.Name -ErrorAction Stop

            foreach ($connector in $connectors) {
                $receiveConnectors += [PSCustomObject]@{
                    Server = $srv.Name
                    Name = $connector.Name
                    Identity = $connector.Identity
                    Bindings = ($connector.Bindings -join "; ")
                    RemoteIPRanges = ($connector.RemoteIPRanges -join "; ")
                    Enabled = $connector.Enabled
                    ProtocolLoggingLevel = $connector.ProtocolLoggingLevel
                    Fqdn = $connector.Fqdn
                    Banner = $connector.Banner
                    MaxMessageSize = $connector.MaxMessageSize
                    MaxRecipientsPerMessage = $connector.MaxRecipientsPerMessage
                    MaxHeaderSize = $connector.MaxHeaderSize
                    MaxHopCount = $connector.MaxHopCount
                    ChunkingEnabled = $connector.ChunkingEnabled
                    AuthMechanism = $connector.AuthMechanism
                    PermissionGroups = $connector.PermissionGroups
                    ConnectionTimeout = $connector.ConnectionTimeout
                    ConnectionInactivityTimeout = $connector.ConnectionInactivityTimeout
                    MessageRateLimit = $connector.MessageRateLimit
                    MessageRateSource = $connector.MessageRateSource
                    RequireTLS = $connector.RequireTLS
                    EnableAuthGSSAPI = $connector.EnableAuthGSSAPI
                    TransportRole = $connector.TransportRole
                }
            }
            Write-Log "  Processed $($connectors.Count) receive connectors from $($srv.Name)" -Level Success
        }
        catch {
            Write-Log "  Error collecting receive connectors from $($srv.Name): $($_.Exception.Message)" -Level Error
        }
    }

    return $receiveConnectors
}

function Get-SendConnectorInfo {
    Write-Log "Collecting Send Connector information..."
    $sendConnectors = @()

    try {
        $connectors = Get-SendConnector -ErrorAction Stop

        foreach ($connector in $connectors) {
            $sendConnectors += [PSCustomObject]@{
                Name = $connector.Name
                Identity = $connector.Identity
                Enabled = $connector.Enabled
                AddressSpaces = ($connector.AddressSpaces -join "; ")
                SourceTransportServers = ($connector.SourceTransportServers -join "; ")
                SmartHosts = ($connector.SmartHosts -join "; ")
                DNSRoutingEnabled = $connector.DNSRoutingEnabled
                UseExternalDNSServersEnabled = $connector.UseExternalDNSServersEnabled
                Fqdn = $connector.Fqdn
                Port = $connector.Port
                ProtocolLoggingLevel = $connector.ProtocolLoggingLevel
                SmartHostAuthMechanism = $connector.SmartHostAuthMechanism
                MaxMessageSize = $connector.MaxMessageSize
                ConnectionInactivityTimeout = $connector.ConnectionInactivityTimeout
                RequireTLS = $connector.RequireTLS
                TlsAuthLevel = $connector.TlsAuthLevel
                TlsDomain = ($connector.TlsDomain -join "; ")
                CloudServicesMailEnabled = $connector.CloudServicesMailEnabled
                IsCoexistenceConnector = $connector.IsCoexistenceConnector
                IsScopedConnector = $connector.IsScopedConnector
                Comment = $connector.Comment
            }
        }
        Write-Log "  Processed $($connectors.Count) send connectors" -Level Success
    }
    catch {
        Write-Log "  Error collecting send connectors: $($_.Exception.Message)" -Level Error
    }

    return $sendConnectors
}

function Get-CertificateInfo {
    param([array]$Servers)

    Write-Log "Collecting Exchange Certificate information..."
    $certificates = @()

    foreach ($srv in $Servers) {
        try {
            $certs = Get-ExchangeCertificate -Server $srv.Name -ErrorAction Stop

            foreach ($cert in $certs) {
                $daysUntilExpiry = if ($cert.NotAfter) {
                    ($cert.NotAfter - (Get-Date)).Days
                } else {
                    $null
                }

                $status = if ($daysUntilExpiry -le 0) {
                    "Expired"
                } elseif ($daysUntilExpiry -le 30) {
                    "Expiring Soon"
                } elseif ($daysUntilExpiry -le 90) {
                    "Warning"
                } else {
                    "Valid"
                }

                $certificates += [PSCustomObject]@{
                    Server = $srv.Name
                    Thumbprint = $cert.Thumbprint
                    Subject = $cert.Subject
                    Issuer = $cert.Issuer
                    FriendlyName = $cert.FriendlyName
                    DomainNames = ($cert.CertificateDomains -join "; ")
                    Services = $cert.Services
                    Status = $cert.Status
                    IsSelfSigned = $cert.IsSelfSigned
                    NotBefore = $cert.NotBefore
                    NotAfter = $cert.NotAfter
                    DaysUntilExpiry = $daysUntilExpiry
                    ExpiryStatus = $status
                    HasPrivateKey = $cert.HasPrivateKey
                    PublicKeySize = $cert.PublicKeySize
                    SignatureAlgorithm = $cert.SignatureAlgorithm
                    SerialNumber = $cert.SerialNumber
                }
            }
            Write-Log "  Processed $($certs.Count) certificates from $($srv.Name)" -Level Success
        }
        catch {
            Write-Log "  Error collecting certificates from $($srv.Name): $($_.Exception.Message)" -Level Error
        }
    }

    return $certificates
}

#endregion

#region Export Functions

function Export-ToCSV {
    param(
        [object]$ReceiveConnectors,
        [object]$SendConnectors,
        [object]$Certificates,
        [string]$Path
    )

    Write-Log "Exporting data to CSV files..."

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

    try {
        if ($ReceiveConnectors.Count -gt 0) {
            $rcFile = Join-Path $Path "ReceiveConnectors_$timestamp.csv"
            $ReceiveConnectors | Export-Csv -Path $rcFile -NoTypeInformation -Encoding UTF8
            Write-Log "  Receive Connectors exported to: $rcFile" -Level Success
        }

        if ($SendConnectors.Count -gt 0) {
            $scFile = Join-Path $Path "SendConnectors_$timestamp.csv"
            $SendConnectors | Export-Csv -Path $scFile -NoTypeInformation -Encoding UTF8
            Write-Log "  Send Connectors exported to: $scFile" -Level Success
        }

        if ($Certificates.Count -gt 0) {
            $certFile = Join-Path $Path "Certificates_$timestamp.csv"
            $Certificates | Export-Csv -Path $certFile -NoTypeInformation -Encoding UTF8
            Write-Log "  Certificates exported to: $certFile" -Level Success
        }
    }
    catch {
        Write-Log "Error exporting to CSV: $($_.Exception.Message)" -Level Error
        throw
    }
}

function Export-ToHTML {
    param(
        [object]$ReceiveConnectors,
        [object]$SendConnectors,
        [object]$Certificates,
        [string]$Path,
        [string]$ServerFilter
    )

    Write-Log "Generating HTML report..."

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportFile = Join-Path $Path "ExchangeConnectorsAndCertificates_$timestamp.html"

    $htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Exchange Connectors and Certificates Report</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 20px;
            background-color: #f5f5f5;
        }
        h1 {
            color: #0078d4;
            border-bottom: 3px solid #0078d4;
            padding-bottom: 10px;
        }
        h2 {
            color: #2b579a;
            margin-top: 30px;
            border-bottom: 2px solid #ddd;
            padding-bottom: 5px;
        }
        .info-box {
            background-color: #e8f4f8;
            border-left: 4px solid #0078d4;
            padding: 15px;
            margin: 20px 0;
        }
        table {
            border-collapse: collapse;
            width: 100%;
            background-color: white;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            margin: 20px 0;
        }
        th {
            background-color: #0078d4;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: bold;
        }
        td {
            padding: 10px;
            border-bottom: 1px solid #ddd;
        }
        tr:hover {
            background-color: #f5f5f5;
        }
        .status-valid { color: green; font-weight: bold; }
        .status-warning { color: orange; font-weight: bold; }
        .status-expiring { color: red; font-weight: bold; }
        .status-expired { color: darkred; font-weight: bold; background-color: #ffeeee; }
        .enabled { color: green; }
        .disabled { color: red; }
        .summary {
            display: inline-block;
            margin: 10px 20px 10px 0;
            padding: 15px 25px;
            background-color: white;
            border-radius: 5px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .summary-label {
            font-size: 14px;
            color: #666;
        }
        .summary-value {
            font-size: 28px;
            font-weight: bold;
            color: #0078d4;
        }
    </style>
</head>
<body>
    <h1>Exchange Connectors and Certificates Report</h1>
    <div class="info-box">
        <strong>Report Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")<br>
        <strong>Target:</strong> $(if($ServerFilter){"Server: $ServerFilter"}else{"All Exchange Servers"})
    </div>
"@

    $htmlBody = ""

    # Summary section
    $htmlBody += "<h2>Summary</h2>"
    $htmlBody += "<div>"
    $htmlBody += "<div class='summary'><div class='summary-label'>Receive Connectors</div><div class='summary-value'>$($ReceiveConnectors.Count)</div></div>"
    $htmlBody += "<div class='summary'><div class='summary-label'>Send Connectors</div><div class='summary-value'>$($SendConnectors.Count)</div></div>"
    $htmlBody += "<div class='summary'><div class='summary-label'>Certificates</div><div class='summary-value'>$($Certificates.Count)</div></div>"

    $expiredCerts = ($Certificates | Where-Object {$_.ExpiryStatus -eq "Expired"}).Count
    $expiringSoonCerts = ($Certificates | Where-Object {$_.ExpiryStatus -eq "Expiring Soon"}).Count

    if ($expiredCerts -gt 0 -or $expiringSoonCerts -gt 0) {
        $htmlBody += "<div class='summary' style='background-color: #fff3cd; border: 2px solid #ffc107;'>"
        $htmlBody += "<div class='summary-label'>Certificates Need Attention</div>"
        $htmlBody += "<div class='summary-value' style='color: #dc3545;'>$($expiredCerts + $expiringSoonCerts)</div></div>"
    }
    $htmlBody += "</div>"

    # Receive Connectors
    if ($ReceiveConnectors.Count -gt 0) {
        $htmlBody += "<h2>Receive Connectors ($($ReceiveConnectors.Count))</h2>"
        $htmlBody += "<table>"
        $htmlBody += "<tr><th>Server</th><th>Name</th><th>Enabled</th><th>Bindings</th><th>FQDN</th><th>Auth Mechanism</th><th>Require TLS</th><th>Max Message Size</th></tr>"

        foreach ($rc in $ReceiveConnectors) {
            $enabledClass = if($rc.Enabled){"enabled"}else{"disabled"}
            $htmlBody += "<tr>"
            $htmlBody += "<td>$($rc.Server)</td>"
            $htmlBody += "<td>$($rc.Name)</td>"
            $htmlBody += "<td class='$enabledClass'>$($rc.Enabled)</td>"
            $htmlBody += "<td>$($rc.Bindings)</td>"
            $htmlBody += "<td>$($rc.Fqdn)</td>"
            $htmlBody += "<td>$($rc.AuthMechanism)</td>"
            $htmlBody += "<td>$($rc.RequireTLS)</td>"
            $htmlBody += "<td>$($rc.MaxMessageSize)</td>"
            $htmlBody += "</tr>"
        }
        $htmlBody += "</table>"
    }

    # Send Connectors
    if ($SendConnectors.Count -gt 0) {
        $htmlBody += "<h2>Send Connectors ($($SendConnectors.Count))</h2>"
        $htmlBody += "<table>"
        $htmlBody += "<tr><th>Name</th><th>Enabled</th><th>Address Spaces</th><th>Source Servers</th><th>Smart Hosts</th><th>DNS Routing</th><th>Require TLS</th><th>Max Message Size</th></tr>"

        foreach ($sc in $SendConnectors) {
            $enabledClass = if($sc.Enabled){"enabled"}else{"disabled"}
            $htmlBody += "<tr>"
            $htmlBody += "<td>$($sc.Name)</td>"
            $htmlBody += "<td class='$enabledClass'>$($sc.Enabled)</td>"
            $htmlBody += "<td>$($sc.AddressSpaces)</td>"
            $htmlBody += "<td>$($sc.SourceTransportServers)</td>"
            $htmlBody += "<td>$($sc.SmartHosts)</td>"
            $htmlBody += "<td>$($sc.DNSRoutingEnabled)</td>"
            $htmlBody += "<td>$($sc.RequireTLS)</td>"
            $htmlBody += "<td>$($sc.MaxMessageSize)</td>"
            $htmlBody += "</tr>"
        }
        $htmlBody += "</table>"
    }

    # Certificates
    if ($Certificates.Count -gt 0) {
        $htmlBody += "<h2>Exchange Certificates ($($Certificates.Count))</h2>"
        $htmlBody += "<table>"
        $htmlBody += "<tr><th>Server</th><th>Subject</th><th>Friendly Name</th><th>Services</th><th>Issuer</th><th>Not After</th><th>Days Until Expiry</th><th>Status</th><th>Self-Signed</th></tr>"

        foreach ($cert in $Certificates | Sort-Object DaysUntilExpiry) {
            $statusClass = switch($cert.ExpiryStatus) {
                "Valid" { "status-valid" }
                "Warning" { "status-warning" }
                "Expiring Soon" { "status-expiring" }
                "Expired" { "status-expired" }
            }

            $htmlBody += "<tr>"
            $htmlBody += "<td>$($cert.Server)</td>"
            $htmlBody += "<td>$($cert.Subject)</td>"
            $htmlBody += "<td>$($cert.FriendlyName)</td>"
            $htmlBody += "<td>$($cert.Services)</td>"
            $htmlBody += "<td>$($cert.Issuer)</td>"
            $htmlBody += "<td>$($cert.NotAfter)</td>"
            $htmlBody += "<td class='$statusClass'>$($cert.DaysUntilExpiry)</td>"
            $htmlBody += "<td class='$statusClass'>$($cert.ExpiryStatus)</td>"
            $htmlBody += "<td>$($cert.IsSelfSigned)</td>"
            $htmlBody += "</tr>"
        }
        $htmlBody += "</table>"
    }

    $htmlFooter = @"
    <div class="info-box" style="margin-top: 40px; background-color: #f8f9fa;">
        <strong>Script:</strong> Get-ExchangeConnectorAndCertificateInfo.ps1<br>
        <strong>Note:</strong> For detailed information, refer to the CSV exports.
    </div>
</body>
</html>
"@

    try {
        $htmlHeader + $htmlBody + $htmlFooter | Out-File -FilePath $reportFile -Encoding UTF8
        Write-Log "  HTML report generated: $reportFile" -Level Success

        # Try to open the report in default browser
        try {
            Start-Process $reportFile -ErrorAction SilentlyContinue
        }
        catch {
            # Silently continue if can't open browser
        }
    }
    catch {
        Write-Log "Error generating HTML report: $($_.Exception.Message)" -Level Error
        throw
    }
}

#endregion

#region Main Script

try {
    Write-Log "======================================" -Level Info
    Write-Log "Exchange Connectors and Certificates Information Gathering" -Level Info
    Write-Log "======================================" -Level Info

    # Validate Exchange Management Shell
    if (-not (Test-ExchangeManagementShell)) {
        Write-Log "Exchange Management Shell not detected. Please run this script from Exchange Management Shell." -Level Error
        exit 1
    }

    # Determine export options
    if (-not $ExportCSV -and -not $ExportHTML -and -not $ExportAll) {
        Write-Log "No export option specified. Defaulting to -ExportAll" -Level Warning
        $ExportAll = $true
    }

    if ($ExportAll) {
        $ExportCSV = $true
        $ExportHTML = $true
    }

    # Validate output path
    if (-not (Test-Path $OutputPath)) {
        Write-Log "Output path does not exist. Creating: $OutputPath"
        New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
    }

    $OutputPath = Resolve-Path $OutputPath
    Write-Log "Output path: $OutputPath"

    # Get Exchange servers
    $exchangeServers = Get-ExchangeServerList -ServerName $Server

    if ($exchangeServers.Count -eq 0) {
        Write-Log "No Exchange servers found." -Level Error
        exit 1
    }

    Write-Log "Found $($exchangeServers.Count) Exchange server(s) to process"

    # Collect data
    $receiveConnectorData = Get-ReceiveConnectorInfo -Servers $exchangeServers
    $sendConnectorData = Get-SendConnectorInfo
    $certificateData = Get-CertificateInfo -Servers $exchangeServers

    # Display summary
    Write-Log "======================================" -Level Info
    Write-Log "Data Collection Summary:" -Level Info
    Write-Log "  Receive Connectors: $($receiveConnectorData.Count)" -Level Success
    Write-Log "  Send Connectors: $($sendConnectorData.Count)" -Level Success
    Write-Log "  Certificates: $($certificateData.Count)" -Level Success

    # Check for certificate issues
    $expiredCerts = $certificateData | Where-Object {$_.ExpiryStatus -eq "Expired"}
    $expiringSoonCerts = $certificateData | Where-Object {$_.ExpiryStatus -eq "Expiring Soon"}

    if ($expiredCerts.Count -gt 0) {
        Write-Log "  WARNING: $($expiredCerts.Count) expired certificate(s) found!" -Level Error
    }
    if ($expiringSoonCerts.Count -gt 0) {
        Write-Log "  WARNING: $($expiringSoonCerts.Count) certificate(s) expiring within 30 days!" -Level Warning
    }

    Write-Log "======================================" -Level Info

    # Export data
    if ($ExportCSV) {
        Export-ToCSV -ReceiveConnectors $receiveConnectorData -SendConnectors $sendConnectorData -Certificates $certificateData -Path $OutputPath
    }

    if ($ExportHTML) {
        Export-ToHTML -ReceiveConnectors $receiveConnectorData -SendConnectors $sendConnectorData -Certificates $certificateData -Path $OutputPath -ServerFilter $Server
    }

    Write-Log "======================================" -Level Info
    Write-Log "Script completed successfully!" -Level Success
    Write-Log "======================================" -Level Info
}
catch {
    Write-Log "Script execution failed: $($_.Exception.Message)" -Level Error
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level Error
    exit 1
}

#endregion
