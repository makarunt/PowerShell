<#
.SYNOPSIS
    Comprehensive Exchange Infrastructure Documentation Script - CORRECTED VERSION

.DESCRIPTION
    This script connects to Exchange On-Premises and/or Exchange Online to generate
    comprehensive documentation including ALL configurations, settings, certificates,
    and infrastructure details. Outputs both CSV and HTML reports for auditing and analysis.

    FIXES APPLIED:
    - Fixed Connect-ExchangeOnline UPN construction bug
    - Completely rewritten CSV export to create usable per-category files
    - Fixed deprecated Get-PublicFolderDatabase cmdlet for Exchange 2013+
    - Optimized mailbox statistics collection (single-pass algorithm)
    - Added HTTPS support for on-premises connections
    - Enhanced error handling and logging
    - Fixed certificate expiry logic
    - Improved input validation
    - Better session cleanup

.PARAMETER Environment
    Specifies the environment to document: OnPremises, Online, or Both

.PARAMETER OutputPath
    Specifies the output directory for reports (default: current directory)

.PARAMETER ExchangeServer
    For on-premises: FQDN of Exchange server to connect to

.PARAMETER Credential
    Credentials for authentication (if not provided, will prompt)

.PARAMETER TenantId
    Azure AD Tenant ID for Exchange Online connection (optional)

.PARAMETER AppId
    Application ID for certificate-based authentication (optional)

.PARAMETER CertificateThumbprint
    Certificate thumbprint for certificate-based authentication (optional)

.PARAMETER IncludeDetailedStats
    Include detailed mailbox and database statistics (may take longer)

.PARAMETER DetailedStatsMailboxLimit
    Maximum number of mailboxes to collect detailed statistics for (default: 100)

.PARAMETER UseHTTPS
    Use HTTPS for on-premises Exchange connection (recommended for security)

.EXAMPLE
    .\Exchange-Documentation-Script-Fixed.ps1 -Environment Both -OutputPath "C:\Reports" -IncludeDetailedStats

.EXAMPLE
    .\Exchange-Documentation-Script-Fixed.ps1 -Environment OnPremises -ExchangeServer "exchange01.contoso.com" -UseHTTPS

.EXAMPLE
    .\Exchange-Documentation-Script-Fixed.ps1 -Environment Online -TenantId "contoso.onmicrosoft.com"

.EXAMPLE
    .\Exchange-Documentation-Script-Fixed.ps1 -Environment Online -AppId "12345678-1234-1234-1234-123456789012" -CertificateThumbprint "ABC123..." -TenantId "contoso.onmicrosoft.com"

.NOTES
    Version: 3.1 (Fixed)
    Author: Exchange Admin Team
    Last Modified: 2026-01-21

    Security Recommendations:
    - For On-Premises: Use Windows Authentication where possible (run as service account)
    - For Online: Use certificate-based authentication for automation
    - Store credentials in Windows Credential Manager, not in scripts
    - Use Read-Only admin roles where possible
    - Always use -UseHTTPS for on-premises connections in production

    Execution Time Estimates:
    - Small environment (<1000 mailboxes): 5-15 minutes
    - Medium environment (1000-10000 mailboxes): 15-45 minutes
    - Large environment (>10000 mailboxes): 45-120 minutes

    Required Permissions:
    - On-Premises: View-Only Organization Management (minimum)
    - Exchange Online: View-Only Organization Management or Global Reader
    - Microsoft Graph: Directory.Read.All, Organization.Read.All
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("OnPremises", "Online", "Both")]
    [string]$Environment,

    [Parameter(Mandatory=$false)]
    [string]$OutputPath = (Get-Location).Path,

    [Parameter(Mandatory=$false)]
    [string]$ExchangeServer,

    [Parameter(Mandatory=$false)]
    [PSCredential]$Credential,

    [Parameter(Mandatory=$false)]
    [string]$TenantId,

    [Parameter(Mandatory=$false)]
    [string]$AppId,

    [Parameter(Mandatory=$false)]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory=$false)]
    [switch]$IncludeDetailedStats,

    [Parameter(Mandatory=$false)]
    [int]$DetailedStatsMailboxLimit = 100,

    [Parameter(Mandatory=$false)]
    [switch]$UseHTTPS
)

# Global variables
$Script:ReportData = @{}
$Script:Timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$Script:HTMLPath = Join-Path $OutputPath "Exchange_Comprehensive_Documentation_$Script:Timestamp.html"
$Script:ConnectedToEXO = $false
$Script:ConnectedToGraph = $false
$Script:ErrorLog = @()

# Function to write progress and log
function Write-LogProgress {
    param(
        [string]$Message,
        [string]$Status = "Processing",
        [switch]$NoProgress
    )

    if (-not $NoProgress) {
        Write-Progress -Activity "Exchange Comprehensive Documentation" -Status $Status -CurrentOperation $Message
    }
    Write-Verbose "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message" -ForegroundColor Green
}

# Function to safely execute commands and handle errors
function Invoke-SafeCommand {
    param(
        [scriptblock]$Command,
        [string]$Description,
        [string]$Category,
        [switch]$Critical
    )

    try {
        Write-LogProgress "Collecting $Description"
        $result = & $Command

        # Store result even if null/empty - let caller decide if that's valid
        $Script:ReportData[$Category] = $result

        if ($null -eq $result) {
            Write-Verbose "No data returned for $Description"
        } elseif ($result -is [Array] -and $result.Count -eq 0) {
            Write-Verbose "Empty collection returned for $Description"
        }

        return $result
    }
    catch {
        $errorMessage = "Failed to collect $Description`: $($_.Exception.Message)"
        $Script:ErrorLog += [PSCustomObject]@{
            Timestamp = Get-Date
            Category = $Category
            Description = $Description
            ErrorMessage = $_.Exception.Message
            FullError = $_.Exception.ToString()
        }

        if ($Critical) {
            Write-Error $errorMessage
            throw
        } else {
            Write-Warning $errorMessage
            Write-Verbose "Full error: $($_.Exception.ToString())"
        }

        return $null
    }
}

# Function to check and install required modules
function Test-RequiredModules {
    param([string]$Environment)

    $modulesNeeded = @()

    if ($Environment -eq "Online" -or $Environment -eq "Both") {
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            $modulesNeeded += "ExchangeOnlineManagement"
        }
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
            Write-Warning "Microsoft.Graph module not found. Some additional data collection will be skipped."
            Write-Warning "To install: Install-Module -Name Microsoft.Graph -Scope CurrentUser"
        }
    }

    if ($modulesNeeded.Count -gt 0) {
        Write-Host "The following modules are required but not installed:" -ForegroundColor Yellow
        $modulesNeeded | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
        Write-Host ""
        Write-Host "To install these modules, run:" -ForegroundColor Cyan
        $modulesNeeded | ForEach-Object { Write-Host "  Install-Module -Name $_ -Scope CurrentUser" -ForegroundColor Cyan }
        Write-Host ""

        $install = Read-Host "Would you like to install these modules now? (Y/N)"
        if ($install -eq "Y" -or $install -eq "y") {
            foreach ($module in $modulesNeeded) {
                try {
                    Write-LogProgress "Installing module: $module"
                    Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
                    Write-Host "Successfully installed $module" -ForegroundColor Green
                }
                catch {
                    Write-Error "Failed to install $module`: $($_.Exception.Message)"
                    return $false
                }
            }
        } else {
            Write-Error "Required modules are not installed. Exiting."
            return $false
        }
    }

    return $true
}

# Function to connect to Exchange On-Premises
function Connect-ExchangeOnPremises {
    param(
        [string]$Server,
        [PSCredential]$Cred,
        [bool]$UseHTTPS
    )

    Write-LogProgress "Connecting to Exchange On-Premises: $Server"

    $protocol = if ($UseHTTPS) { "https" } else { "http" }
    $uri = "$protocol`://$Server/PowerShell/"

    if (-not $UseHTTPS) {
        Write-Warning "Using HTTP connection. Consider using -UseHTTPS for secure connections in production."
    }

    try {
        $sessionParams = @{
            ConfigurationName = "Microsoft.Exchange"
            ConnectionUri = $uri
            Authentication = "Kerberos"
        }

        if ($Cred) {
            $sessionParams.Credential = $Cred
        }

        $Session = New-PSSession @sessionParams

        if (-not $Session) {
            throw "Failed to create PowerShell session"
        }

        $importResult = Import-PSSession $Session -DisableNameChecking -AllowClobber -WarningAction SilentlyContinue -ErrorAction Stop

        if (-not $importResult) {
            throw "Failed to import Exchange cmdlets from session"
        }

        # Test connection
        $null = Get-OrganizationConfig -ErrorAction Stop

        Write-LogProgress "Successfully connected to Exchange On-Premises"
        return $Session
    }
    catch {
        Write-Error "Failed to connect to Exchange On-Premises: $($_.Exception.Message)"
        if ($Session) {
            Remove-PSSession $Session -ErrorAction SilentlyContinue
        }
        return $null
    }
}

# Function to connect to Exchange Online
function Connect-ExchangeOnline {
    param(
        [string]$TenantId,
        [string]$AppId,
        [string]$CertThumbprint
    )

    Write-LogProgress "Connecting to Exchange Online"

    try {
        Import-Module ExchangeOnlineManagement -Force -ErrorAction Stop

        # Determine connection method
        if ($AppId -and $CertThumbprint -and $TenantId) {
            # Certificate-based authentication
            Write-LogProgress "Using certificate-based authentication"
            Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $CertThumbprint -Organization $TenantId -ShowProgress $false -ErrorAction Stop
        } elseif ($TenantId) {
            # Interactive authentication with specific tenant
            Write-LogProgress "Using interactive authentication with tenant: $TenantId"
            # FIXED: Use DelegatedOrganization instead of constructing invalid UPN
            Connect-ExchangeOnline -DelegatedOrganization $TenantId -ShowProgress $false -ErrorAction Stop
        } else {
            # Standard interactive authentication
            Write-LogProgress "Using interactive authentication"
            Connect-ExchangeOnline -ShowProgress $false -ErrorAction Stop
        }

        # Test connection
        $null = Get-OrganizationConfig -ErrorAction Stop
        $Script:ConnectedToEXO = $true

        Write-LogProgress "Successfully connected to Exchange Online"
        return $true
    }
    catch {
        Write-Error "Failed to connect to Exchange Online: $($_.Exception.Message)"
        return $false
    }
}

# Function to connect to Microsoft Graph
function Connect-MicrosoftGraph {
    param([string]$TenantId)

    try {
        if (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication) {
            Write-LogProgress "Connecting to Microsoft Graph"
            Import-Module Microsoft.Graph.Authentication -Force -ErrorAction Stop

            # Reduced scope - removed SecurityEvents.Read.All which requires high admin consent
            $scopes = @(
                "Directory.Read.All",
                "Organization.Read.All",
                "Policy.Read.All"
            )

            $connectParams = @{
                Scopes = $scopes
                NoWelcome = $true
                ErrorAction = 'Stop'
            }

            if ($TenantId) {
                $connectParams.TenantId = $TenantId
            }

            Connect-MgGraph @connectParams

            $Script:ConnectedToGraph = $true
            Write-LogProgress "Successfully connected to Microsoft Graph"
            return $true
        } else {
            Write-Verbose "Microsoft.Graph module not available"
            return $false
        }
    }
    catch {
        Write-Warning "Could not connect to Microsoft Graph: $($_.Exception.Message)"
        Write-Verbose "Graph connection is optional. Continuing without Graph data."
        return $false
    }
}

# Function to collect Exchange On-Premises data
function Get-ExchangeOnPremisesData {
    Write-LogProgress "Starting comprehensive Exchange On-Premises data collection"

    # Organization Configuration
    Invoke-SafeCommand -Command {
        Get-OrganizationConfig | Select-Object Name, ExchangeVersion, AdminDisplayVersion, IsDehydrated,
        HybridConfigurationStatus, MaxReceiveSize, MaxSendSize, DefaultPublicFolderDatabase,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Organization Configuration" -Category "OrganizationConfig"

    # Exchange Servers with detailed information
    Invoke-SafeCommand -Command {
        Get-ExchangeServer | Select-Object Name, ServerRole, AdminDisplayVersion, Edition, FQDN, Site,
        IsHubTransportServer, IsClientAccessServer, IsMailboxServer, IsUnifiedMessagingServer, IsEdgeServer,
        NetworkAddress, OrganizationalUnit, WhenCreated, WhenChanged,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Exchange Servers" -Category "ExchangeServers"

    # Exchange Certificates - CRITICAL for SMTP, EWS, etc.
    Invoke-SafeCommand -Command {
        $certs = @()

        # Try to get servers first
        try {
            $servers = Get-ExchangeServer -ErrorAction Stop
        }
        catch {
            Write-Warning "Could not retrieve Exchange servers list: $($_.Exception.Message)"
            return $null
        }

        foreach ($server in $servers) {
            try {
                $serverCerts = Get-ExchangeCertificate -Server $server.Name -ErrorAction Stop | Select-Object `
                    @{N='Server';E={$server.Name}}, `
                    Thumbprint, Subject, Issuer, NotBefore, NotAfter, Status, `
                    Services, CertificateDomains, IsSelfSigned, HasPrivateKey, `
                    @{N='DaysUntilExpiry';E={
                        if ($_.NotAfter) {
                            [Math]::Round(($_.NotAfter - (Get-Date)).TotalDays, 0)
                        } else {
                            $null
                        }
                    }}, `
                    @{N='IsExpired';E={
                        if ($_.NotAfter) {
                            $_.NotAfter -lt (Get-Date)
                        } else {
                            $false
                        }
                    }}, `
                    @{N='CollectedDate';E={Get-Date}}
                $certs += $serverCerts
            }
            catch {
                Write-Warning "Could not retrieve certificates from server $($server.Name): $($_.Exception.Message)"
            }
        }
        return $certs
    } -Description "Exchange Certificates" -Category "ExchangeCertificates"

    # Database Information with detailed settings
    Invoke-SafeCommand -Command {
        Get-MailboxDatabase | Select-Object Name, Server, MasterServerOrAvailabilityGroup, EdbFilePath,
        LogFolderPath, CircularLoggingEnabled, MaintenanceSchedule, QuotaNotificationSchedule,
        ProhibitSendQuota, ProhibitSendReceiveQuota, IssueWarningQuota, DeletedItemRetention,
        MailboxRetention, IndexEnabled, BackgroundDatabaseMaintenance, AllowFileRestore,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Mailbox Databases" -Category "MailboxDatabases"

    # Database Copies and Health
    Invoke-SafeCommand -Command {
        Get-MailboxDatabaseCopyStatus -ErrorAction SilentlyContinue | Select-Object Name, Status, CopyQueueLength, ReplayQueueLength,
        LastInspectedLogTime, ContentIndexState, ActivationSuspended,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Database Copy Status" -Category "DatabaseCopyStatus"

    # Public Folder Databases/Mailboxes - FIXED for Exchange 2013+
    Invoke-SafeCommand -Command {
        try {
            # Try old cmdlet for Exchange 2010 and earlier
            $pfDatabases = Get-PublicFolderDatabase -ErrorAction Stop | Select-Object Name, Server, EdbFilePath,
            LogFolderPath, MaintenanceSchedule, MaxItemSize, ProhibitPostQuota, IssueWarningQuota,
            @{N='Type';E={'Legacy Database'}},
            @{N='CollectedDate';E={Get-Date}}

            if ($pfDatabases) {
                return $pfDatabases
            }
        }
        catch {
            Write-Verbose "Legacy public folder databases not found or not supported. Checking for modern public folder mailboxes."
        }

        # Get modern public folder mailboxes (Exchange 2013+)
        try {
            $pfMailboxes = Get-Mailbox -PublicFolder -ErrorAction Stop | Select-Object Name, Database,
            ProhibitSendQuota, IssueWarningQuota, IsRootPublicFolderMailbox,
            @{N='Type';E={'Modern Public Folder Mailbox'}},
            @{N='CollectedDate';E={Get-Date}}

            return $pfMailboxes
        }
        catch {
            Write-Verbose "No public folder mailboxes found: $($_.Exception.Message)"
            return $null
        }
    } -Description "Public Folder Databases/Mailboxes" -Category "PublicFolderDatabases"

    # Database Availability Groups with detailed configuration
    Invoke-SafeCommand -Command {
        Get-DatabaseAvailabilityGroup -ErrorAction SilentlyContinue | Select-Object Name, Servers,
        WitnessServer, WitnessDirectory, AlternateWitnessServer, NetworkCompression, NetworkEncryption,
        ReplicationPort, DatacenterActivationMode, ThirdPartyReplication,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Database Availability Groups" -Category "DatabaseAvailabilityGroups"

    # Receive Connectors - Including SMTP Relay configurations
    Invoke-SafeCommand -Command {
        Get-ReceiveConnector | Select-Object Identity, Server, Bindings, RemoteIPRanges, AuthMechanism,
        PermissionGroups, MaxMessageSize, ConnectionTimeout, MaxInboundConnection, RequireTLS,
        EnableAuthGSSAPI, ExtendedProtectionPolicy, SuppressXAnonymousTls, AdvertiseClientSettings,
        Banner, Comment, Enabled, Fqdn, LongAddressesEnabled, OrarEnabled, PipeliningEnabled,
        ProtocolLoggingLevel, SizeEnabled, TarpitInterval, TransportRole,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Receive Connectors (SMTP Relay)" -Category "ReceiveConnectors"

    # Send Connectors - Including SMTP Relay configurations
    Invoke-SafeCommand -Command {
        Get-SendConnector | Select-Object Identity, AddressSpaces, SourceTransportServers, SmartHosts,
        Port, RequireTLS, SmartHostAuthMechanism, UseExternalDNSServersEnabled, MaxMessageSize,
        ConnectionInactivityTimeout, DnsRoutingEnabled, ErrorPolicies, ForceHELO, Fqdn,
        IgnoreSTARTTLS, IsScopedConnector, IsSmtpConnector, LinkedReceiveConnector, ProtocolLoggingLevel,
        SmartHostsString, TlsAuthLevel, TlsCertificateName, TlsDomain,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Send Connectors (SMTP Relay)" -Category "SendConnectors"

    # Transport Configuration
    Invoke-SafeCommand -Command {
        Get-TransportConfig | Select-Object MaxDumpsterSizePerDatabase, MaxDumpsterTime,
        MaxReceiveSize, MaxSendSize, ExternalPostmasterAddress, GenerateCopyOfDSNFor,
        InternalSMTPServers, JournalingReportNdrTo, MaxRecipientEnvelopeLimit,
        OrganizationFederatedMailbox, RedirectUnprovisionedUserMessagesTo, ShadowRedundancyEnabled,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Transport Configuration" -Category "TransportConfiguration"

    # Transport Rules with detailed conditions and actions
    Invoke-SafeCommand -Command {
        Get-TransportRule | Select-Object Name, Priority, State, Mode, Description, Conditions, Actions,
        Exceptions, Comments, RuleVersion, WhenChanged,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Transport Rules" -Category "TransportRules"

    # Accepted Domains
    Invoke-SafeCommand -Command {
        Get-AcceptedDomain | Select-Object Name, DomainName, DomainType, Default, MatchSubDomains,
        AddressBookEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "Accepted Domains" -Category "AcceptedDomains"

    # Remote Domains
    Invoke-SafeCommand -Command {
        Get-RemoteDomain | Select-Object Name, DomainName, AllowedOOFType, AutoReplyEnabled,
        AutoForwardEnabled, DeliveryReportEnabled, NDREnabled, MeetingForwardNotificationEnabled,
        UseSimpleDisplayName, @{N='CollectedDate';E={Get-Date}}
    } -Description "Remote Domains" -Category "RemoteDomains"

    # Email Address Policies
    Invoke-SafeCommand -Command {
        Get-EmailAddressPolicy | Select-Object Name, Priority, EnabledEmailAddressTemplates,
        RecipientFilter, RecipientContainer, @{N='CollectedDate';E={Get-Date}}
    } -Description "Email Address Policies" -Category "EmailAddressPolicies"

    # Virtual Directories - Critical for client connectivity
    Invoke-SafeCommand -Command {
        $vdirs = @()

        # OWA Virtual Directories
        try {
            $vdirs += Get-OwaVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl,
                @{N='Type';E={'OWA'}}, DefaultDomain, LogonFormat, ClientAuthCleanupLevel,
                ExternalAuthenticationMethods, InternalAuthenticationMethods, WindowsAuthentication,
                @{N='CollectedDate';E={Get-Date}}
        } catch { Write-Verbose "Could not get OWA virtual directories" }

        # ECP Virtual Directories
        try {
            $vdirs += Get-EcpVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl,
                @{N='Type';E={'ECP'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
                @{N='CollectedDate';E={Get-Date}}
        } catch { Write-Verbose "Could not get ECP virtual directories" }

        # ActiveSync Virtual Directories
        try {
            $vdirs += Get-ActiveSyncVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl,
                @{N='Type';E={'ActiveSync'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
                ClientCertAuth, CompressionEnabled, WindowsAuthEnabled,
                @{N='CollectedDate';E={Get-Date}}
        } catch { Write-Verbose "Could not get ActiveSync virtual directories" }

        # EWS Virtual Directories - CRITICAL
        try {
            $vdirs += Get-WebServicesVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl,
                @{N='Type';E={'EWS'}}, CertificateAuthentication, WSSecurityAuthentication, OAuthAuthentication,
                ExternalAuthenticationMethods, InternalAuthenticationMethods, WindowsAuthentication,
                @{N='CollectedDate';E={Get-Date}}
        } catch { Write-Verbose "Could not get EWS virtual directories" }

        # OAB Virtual Directories
        try {
            $vdirs += Get-OabVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl,
                @{N='Type';E={'OAB'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
                RequireSSL, @{N='CollectedDate';E={Get-Date}}
        } catch { Write-Verbose "Could not get OAB virtual directories" }

        # Autodiscover Virtual Directories
        try {
            $vdirs += Get-AutodiscoverVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl,
                @{N='Type';E={'Autodiscover'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
                WindowsAuthentication, WSSecurityAuthentication,
                @{N='CollectedDate';E={Get-Date}}
        } catch { Write-Verbose "Could not get Autodiscover virtual directories" }

        # MAPI Virtual Directories
        try {
            $vdirs += Get-MapiVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl,
                @{N='Type';E={'MAPI'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
                @{N='CollectedDate';E={Get-Date}}
        } catch { Write-Verbose "Could not get MAPI virtual directories" }

        # PowerShell Virtual Directories
        try {
            $vdirs += Get-PowerShellVirtualDirectory -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalUrl, ExternalUrl,
                @{N='Type';E={'PowerShell'}}, ExternalAuthenticationMethods, InternalAuthenticationMethods,
                RequireSSL, CertificateAuthentication,
                @{N='CollectedDate';E={Get-Date}}
        } catch { Write-Verbose "Could not get PowerShell virtual directories" }

        return $vdirs
    } -Description "Virtual Directories (OWA, EWS, ActiveSync, etc.)" -Category "VirtualDirectories"

    # Client Access Services
    Invoke-SafeCommand -Command {
        Get-ClientAccessService -ErrorAction SilentlyContinue | Select-Object Name, Server, AutoDiscoverServiceInternalUri,
        AutoDiscoverSiteScope, AlternateServiceAccountConfiguration,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Client Access Services" -Category "ClientAccessServices"

    # Outlook Anywhere Configuration
    Invoke-SafeCommand -Command {
        Get-OutlookAnywhere -ErrorAction SilentlyContinue | Select-Object Identity, Server, InternalHostname, ExternalHostname,
        InternalClientAuthenticationMethod, ExternalClientAuthenticationMethod, IISAuthenticationMethods,
        SSLOffloading, ExternalClientsRequireSsl, InternalClientsRequireSsl,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Outlook Anywhere (RPC over HTTP)" -Category "OutlookAnywhere"

    # Federation Trust and Organization Relationships
    Invoke-SafeCommand -Command {
        Get-FederationTrust -ErrorAction SilentlyContinue | Select-Object Name, ApplicationUri,
        TokenIssuerUris, OrgCertificate, TokenIssuerCertificate, TokenIssuerPrevCertificate,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Federation Trust" -Category "FederationTrust"

    Invoke-SafeCommand -Command {
        Get-OrganizationRelationship -ErrorAction SilentlyContinue | Select-Object Name, DomainNames,
        FreeBusyAccessEnabled, FreeBusyAccessLevel, FreeBusyAccessScope, MailboxMoveEnabled,
        DeliveryReportEnabled, MailTipsAccessEnabled, MailTipsAccessLevel, MailTipsAccessScope,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Organization Relationships" -Category "OrganizationRelationships"

    # Sharing Policies
    Invoke-SafeCommand -Command {
        Get-SharingPolicy -ErrorAction SilentlyContinue | Select-Object Name, Domains, Enabled, Default,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Sharing Policies" -Category "SharingPolicies"

    # Retention Policies and Tags
    Invoke-SafeCommand -Command {
        Get-RetentionPolicy -ErrorAction SilentlyContinue | Select-Object Name, RetentionPolicyTagLinks, IsDefault,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Retention Policies" -Category "RetentionPolicies"

    Invoke-SafeCommand -Command {
        Get-RetentionPolicyTag -ErrorAction SilentlyContinue | Select-Object Name, Type, RetentionEnabled, AgeLimitForRetention,
        RetentionAction, MessageClass, @{N='CollectedDate';E={Get-Date}}
    } -Description "Retention Policy Tags" -Category "RetentionPolicyTags"

    # Address Lists and Global Address Lists
    Invoke-SafeCommand -Command {
        Get-AddressList -ErrorAction SilentlyContinue | Select-Object Name, RecipientFilter, RecipientContainer, DisplayName,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Address Lists" -Category "AddressLists"

    Invoke-SafeCommand -Command {
        Get-GlobalAddressList -ErrorAction SilentlyContinue | Select-Object Name, RecipientFilter, RecipientContainer,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Global Address Lists" -Category "GlobalAddressLists"

    # Offline Address Books
    Invoke-SafeCommand -Command {
        Get-OfflineAddressBook -ErrorAction SilentlyContinue | Select-Object Name, AddressLists, Server, PublicFolderDatabase,
        Schedule, IsDefault, @{N='CollectedDate';E={Get-Date}}
    } -Description "Offline Address Books" -Category "OfflineAddressBooks"

    # Mailbox Statistics Summary - OPTIMIZED with single-pass algorithm
    Invoke-SafeCommand -Command {
        $mailboxes = Get-Mailbox -ResultSize Unlimited

        # Single pass through collection instead of 7 separate Where-Object calls
        $stats = @{
            TotalMailboxes = 0
            UserMailboxes = 0
            SharedMailboxes = 0
            ResourceMailboxes = 0
            RoomMailboxes = 0
            EquipmentMailboxes = 0
            LinkedMailboxes = 0
            CollectedDate = Get-Date
        }

        foreach ($mailbox in $mailboxes) {
            $stats.TotalMailboxes++
            switch ($mailbox.RecipientTypeDetails) {
                'UserMailbox' { $stats.UserMailboxes++ }
                'SharedMailbox' { $stats.SharedMailboxes++ }
                'RoomMailbox' { $stats.RoomMailboxes++ }
                'EquipmentMailbox' { $stats.EquipmentMailboxes++ }
                'LinkedMailbox' { $stats.LinkedMailboxes++ }
                default {
                    if ($_ -like '*Resource*') {
                        $stats.ResourceMailboxes++
                    }
                }
            }
        }

        return [PSCustomObject]$stats
    } -Description "Mailbox Statistics" -Category "MailboxStatistics"

    # Detailed Mailbox Statistics (if requested)
    if ($IncludeDetailedStats) {
        Invoke-SafeCommand -Command {
            Write-LogProgress "Collecting detailed mailbox statistics (sampling $DetailedStatsMailboxLimit mailboxes)"
            $mailboxStats = @()
            $mailboxes = Get-Mailbox -ResultSize $DetailedStatsMailboxLimit

            foreach ($mailbox in $mailboxes) {
                try {
                    $stats = Get-MailboxStatistics $mailbox.Identity -ErrorAction Stop
                    $mailboxStats += [PSCustomObject]@{
                        DisplayName = $mailbox.DisplayName
                        PrimarySmtpAddress = $mailbox.PrimarySmtpAddress
                        RecipientTypeDetails = $mailbox.RecipientTypeDetails
                        Database = $stats.Database
                        TotalItemSize = $stats.TotalItemSize
                        ItemCount = $stats.ItemCount
                        DeletedItemCount = $stats.DeletedItemCount
                        LastLogonTime = $stats.LastLogonTime
                        LastLoggedOnUserAccount = $stats.LastLoggedOnUserAccount
                        CollectedDate = Get-Date
                    }
                }
                catch {
                    Write-Verbose "Could not get statistics for mailbox $($mailbox.DisplayName): $($_.Exception.Message)"
                }
            }
            return $mailboxStats
        } -Description "Detailed Mailbox Statistics (Sample of $DetailedStatsMailboxLimit)" -Category "DetailedMailboxStats"
    }

    # Hybrid Configuration (if exists)
    Invoke-SafeCommand -Command {
        Get-HybridConfiguration -ErrorAction SilentlyContinue | Select-Object Identity,
        OnPremisesSmartHost, Domains, Features, TlsCertificateName, EdgeTransportServers,
        ReceivingTransportServers, SendingTransportServers, ClientAccessServers,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Hybrid Configuration" -Category "HybridConfiguration"

    # Edge Synchronization (if applicable)
    Invoke-SafeCommand -Command {
        Get-EdgeSubscription -ErrorAction SilentlyContinue | Select-Object Name, Site, Domain,
        CreateDate, @{N='CollectedDate';E={Get-Date}}
    } -Description "Edge Subscriptions" -Category "EdgeSubscriptions"

    # Message Classifications
    Invoke-SafeCommand -Command {
        Get-MessageClassification -ErrorAction SilentlyContinue | Select-Object Name, DisplayName,
        SenderDescription, RecipientDescription, ClassificationID,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Message Classifications" -Category "MessageClassifications"

    # Throttling Policies
    Invoke-SafeCommand -Command {
        Get-ThrottlingPolicy -ErrorAction SilentlyContinue | Select-Object Name, IsDefault, AnonymousMaxConcurrency,
        EasMaxConcurrency, EwsMaxConcurrency, ImapMaxConcurrency, OutlookServiceMaxConcurrency,
        OwaMaxConcurrency, PopMaxConcurrency, PowerShellMaxConcurrency, RcaMaxConcurrency,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Throttling Policies" -Category "ThrottlingPolicies"

    # Mobile Device Mailbox Policies
    Invoke-SafeCommand -Command {
        Get-ActiveSyncMailboxPolicy -ErrorAction SilentlyContinue | Select-Object Name, AllowNonProvisionableDevices,
        PasswordEnabled, AlphanumericPasswordRequired, PasswordRecoveryEnabled, DeviceEncryptionEnabled,
        AttachmentsEnabled, MaxAttachmentSize, AllowStorageCard, AllowCamera, AllowWiFi, AllowBluetooth,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Mobile Device Mailbox Policies" -Category "MobileDeviceMailboxPolicies"

    # OWA Mailbox Policies
    Invoke-SafeCommand -Command {
        Get-OwaMailboxPolicy -ErrorAction SilentlyContinue | Select-Object Name, DirectFileAccessOnPublicComputersEnabled,
        DirectFileAccessOnPrivateComputersEnabled, WebReadyDocumentViewingOnPublicComputersEnabled,
        ForceWebReadyDocumentViewingFirstOnPublicComputers, ActiveSyncIntegrationEnabled,
        AllowOfflineOn, ExternalImageProxyEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "OWA Mailbox Policies" -Category "OWAMailboxPolicies"

    # Journal Rules
    Invoke-SafeCommand -Command {
        Get-JournalRule -ErrorAction SilentlyContinue | Select-Object Name, JournalEmailAddress, Scope, Recipient, Enabled,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Journal Rules" -Category "JournalRules"

    # Management Role Assignments (Security)
    Invoke-SafeCommand -Command {
        Get-ManagementRoleAssignment -ErrorAction SilentlyContinue | Select-Object Name, Role, RoleAssignee, RoleAssigneeType,
        AssignmentMethod, IsValid, @{N='CollectedDate';E={Get-Date}}
    } -Description "Management Role Assignments" -Category "ManagementRoleAssignments"

    Write-LogProgress "Completed comprehensive Exchange On-Premises data collection"
}

# Function to collect Exchange Online data
function Get-ExchangeOnlineData {
    Write-LogProgress "Starting comprehensive Exchange Online data collection"

    # Organization Configuration
    Invoke-SafeCommand -Command {
        Get-OrganizationConfig | Select-Object Name, DisplayName, DefaultMailTip, MailTipsAllTipsEnabled,
        MailTipsExternalRecipientsTipsEnabled, MailTipsGroupMetricsEnabled, MailTipsLargeAudienceThreshold,
        IsDehydrated, HybridConfigurationStatus, MaxReceiveSize, MaxSendSize,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Organization Configuration" -Category "EXO_OrganizationConfig"

    # Tenant Information
    Invoke-SafeCommand -Command {
        $tenant = Get-OrganizationConfig
        $defaultDomain = Get-AcceptedDomain | Where-Object {$_.Default -eq $true}
        $tenantInfo = @{
            TenantName = $tenant.Name
            DisplayName = $tenant.DisplayName
            DefaultDomain = $defaultDomain.DomainName
            TotalDomains = (Get-AcceptedDomain).Count
            ExchangeVersion = $tenant.AdminDisplayVersion
            IsDehydrated = $tenant.IsDehydrated
            HybridConfigurationStatus = $tenant.HybridConfigurationStatus
            CollectedDate = Get-Date
        }
        return [PSCustomObject]$tenantInfo
    } -Description "Tenant Information" -Category "EXO_TenantInfo"

    # Mailbox Plans
    Invoke-SafeCommand -Command {
        Get-MailboxPlan -ErrorAction SilentlyContinue | Select-Object DisplayName, MaxSendSize, MaxReceiveSize, ProhibitSendQuota,
        ProhibitSendReceiveQuota, IssueWarningQuota, RetainDeletedItemsFor, RoleAssignmentPolicy,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Mailbox Plans" -Category "EXO_MailboxPlans"

    # Accepted Domains
    Invoke-SafeCommand -Command {
        Get-AcceptedDomain | Select-Object Name, DomainName, DomainType, Default, MatchSubDomains,
        AddressBookEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "Accepted Domains" -Category "EXO_AcceptedDomains"

    # Remote Domains
    Invoke-SafeCommand -Command {
        Get-RemoteDomain | Select-Object Name, DomainName, AllowedOOFType, AutoReplyEnabled,
        AutoForwardEnabled, DeliveryReportEnabled, NDREnabled, MeetingForwardNotificationEnabled,
        UseSimpleDisplayName, @{N='CollectedDate';E={Get-Date}}
    } -Description "Remote Domains" -Category "EXO_RemoteDomains"

    # Transport Configuration
    Invoke-SafeCommand -Command {
        Get-TransportConfig | Select-Object MaxReceiveSize, MaxSendSize, ExternalPostmasterAddress,
        GenerateCopyOfDSNFor, JournalingReportNdrTo, MaxRecipientEnvelopeLimit,
        OrganizationFederatedMailbox, RedirectUnprovisionedUserMessagesTo,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Transport Configuration" -Category "EXO_TransportConfiguration"

    # Transport Rules
    Invoke-SafeCommand -Command {
        Get-TransportRule | Select-Object Name, Priority, State, Mode, Description, Comments,
        RuleVersion, WhenChanged, @{N='CollectedDate';E={Get-Date}}
    } -Description "Transport Rules" -Category "EXO_TransportRules"

    # Inbound Connectors - SMTP Relay for Exchange Online
    Invoke-SafeCommand -Command {
        Get-InboundConnector | Select-Object Name, ConnectorType, ConnectorSource, SenderDomains,
        SenderIPAddresses, RequireTls, RestrictDomainsToIPAddresses, Enabled, Comment,
        TlsSenderCertificateName, CloudServicesMailEnabled, TreatMessagesAsInternal,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Inbound Connectors (SMTP Relay)" -Category "EXO_InboundConnectors"

    # Outbound Connectors - SMTP Relay for Exchange Online
    Invoke-SafeCommand -Command {
        Get-OutboundConnector | Select-Object Name, ConnectorType, RecipientDomains, SmartHosts,
        TlsDomain, UseMxRecord, RouteAllMessagesViaOnPremises, Enabled, Comment,
        TlsSettings, IsTransportRuleScoped, CloudServicesMailEnabled, AllAcceptedDomains,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Outbound Connectors (SMTP Relay)" -Category "EXO_OutboundConnectors"

    # Anti-Spam Policies (Exchange Online Protection)
    Invoke-SafeCommand -Command {
        Get-HostedContentFilterPolicy | Select-Object Name, SpamAction, HighConfidenceSpamAction,
        PhishSpamAction, BulkSpamAction, QuarantineRetentionPeriod, EndUserSpamNotificationFrequency,
        TestModeAction, IncreaseScoreWithImageLinks, IncreaseScoreWithNumericIps,
        IncreaseScoreWithRedirectToOtherPort, IncreaseScoreWithBizOrInfoUrls,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Anti-Spam Policies (EOP)" -Category "EXO_AntiSpamPolicies"

    # Anti-Malware Policies (Exchange Online Protection)
    Invoke-SafeCommand -Command {
        Get-MalwareFilterPolicy | Select-Object Name, Action, EnableFileFilter, FileTypes,
        EnableInternalSenderAdminNotifications, EnableInternalSenderNotifications,
        InternalSenderAdminAddress, CustomNotifications, CustomFromAddress, CustomFromName,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Anti-Malware Policies (EOP)" -Category "EXO_AntiMalwarePolicies"

    # Connection Filter Policies (IP Allow/Block Lists)
    Invoke-SafeCommand -Command {
        Get-HostedConnectionFilterPolicy | Select-Object Name, IPAllowList, IPBlockList,
        EnableSafeList, DirectoryBasedEdgeBlockMode, @{N='CollectedDate';E={Get-Date}}
    } -Description "Connection Filter Policies (IP Lists)" -Category "EXO_ConnectionFilterPolicies"

    # Safe Attachments Policies (Defender for Office 365)
    Invoke-SafeCommand -Command {
        Get-SafeAttachmentPolicy -ErrorAction SilentlyContinue | Select-Object Name, Enable, Action,
        Redirect, RedirectAddress, ActionOnError, @{N='CollectedDate';E={Get-Date}}
    } -Description "Safe Attachments Policies (Defender)" -Category "EXO_SafeAttachmentPolicies"

    # Safe Links Policies (Defender for Office 365)
    Invoke-SafeCommand -Command {
        Get-SafeLinksPolicy -ErrorAction SilentlyContinue | Select-Object Name, IsEnabled,
        ScanUrls, EnableForInternalSenders, TrackClicks, AllowClickThrough, EnableSafeLinksForTeams,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Safe Links Policies (Defender)" -Category "EXO_SafeLinksPolicies"

    # ATP Policies (Defender for Office 365)
    Invoke-SafeCommand -Command {
        Get-AtpPolicyForO365 -ErrorAction SilentlyContinue | Select-Object Name, EnableATPForSPOTeamsODB,
        EnableSafeDocs, AllowSafeDocsOpen, @{N='CollectedDate';E={Get-Date}}
    } -Description "ATP Policies (Defender)" -Category "EXO_ATPPolicies"

    # Anti-Phishing Policies
    Invoke-SafeCommand -Command {
        Get-AntiPhishPolicy -ErrorAction SilentlyContinue | Select-Object Name, Enabled,
        EnableMailboxIntelligence, EnableMailboxIntelligenceProtection, EnableSpoofIntelligence,
        EnableFirstContactSafetyTips, EnableSimilarUsersSafetyTips, EnableSimilarDomainsSafetyTips,
        EnableUnusualCharactersSafetyTips, @{N='CollectedDate';E={Get-Date}}
    } -Description "Anti-Phishing Policies" -Category "EXO_AntiPhishingPolicies"

    # DKIM Configuration
    Invoke-SafeCommand -Command {
        $domains = Get-AcceptedDomain | Where-Object {$_.DomainType -ne "InternalRelay"}
        $dkimConfig = @()
        foreach ($domain in $domains) {
            try {
                $dkim = Get-DkimSigningConfig -Identity $domain.DomainName -ErrorAction SilentlyContinue
                if ($dkim) {
                    $dkimConfig += $dkim | Select-Object Domain, Enabled, Status, Selector1CNAME,
                        Selector2CNAME, @{N='CollectedDate';E={Get-Date}}
                }
            }
            catch {
                Write-Verbose "Domain $($domain.DomainName) does not support DKIM or is not configured: $($_.Exception.Message)"
            }
        }
        return $dkimConfig
    } -Description "DKIM Configuration" -Category "EXO_DKIMConfiguration"

    # Detailed Mailbox Statistics - OPTIMIZED with single-pass algorithm
    Invoke-SafeCommand -Command {
        Write-LogProgress "Collecting mailbox statistics (this may take a while for large tenants)"
        $mailboxes = Get-EXOMailbox -ResultSize Unlimited -PropertySets Minimum

        # Single pass through collection instead of 8 separate Where-Object calls
        $stats = @{
            TotalMailboxes = 0
            UserMailboxes = 0
            SharedMailboxes = 0
            ResourceMailboxes = 0
            EquipmentMailboxes = 0
            RoomMailboxes = 0
            DiscoveryMailboxes = 0
            GroupMailboxes = 0
            CollectedDate = Get-Date
        }

        foreach ($mailbox in $mailboxes) {
            $stats.TotalMailboxes++
            switch ($mailbox.RecipientTypeDetails) {
                'UserMailbox' { $stats.UserMailboxes++ }
                'SharedMailbox' { $stats.SharedMailboxes++ }
                'EquipmentMailbox' { $stats.EquipmentMailboxes++ }
                'RoomMailbox' { $stats.RoomMailboxes++ }
                'DiscoveryMailbox' { $stats.DiscoveryMailboxes++ }
                'GroupMailbox' { $stats.GroupMailboxes++ }
                default {
                    if ($_ -like '*Resource*') {
                        $stats.ResourceMailboxes++
                    }
                }
            }
        }

        return [PSCustomObject]$stats
    } -Description "Mailbox Statistics" -Category "EXO_MailboxStatistics"

    # Distribution Groups Statistics - OPTIMIZED
    Invoke-SafeCommand -Command {
        $groups = Get-DistributionGroup -ResultSize Unlimited
        $dynamicGroups = Get-DynamicDistributionGroup -ResultSize Unlimited

        $groupStats = @{
            TotalDistributionGroups = $groups.Count
            SecurityGroups = ($groups | Where-Object {$_.GroupType -like '*Security*'}).Count
            UniversalGroups = ($groups | Where-Object {$_.GroupType -like '*Universal*'}).Count
            DynamicGroups = $dynamicGroups.Count
            CollectedDate = Get-Date
        }
        return [PSCustomObject]$groupStats
    } -Description "Distribution Group Statistics" -Category "EXO_DistributionGroupStats"

    # Mobile Device Access Rules
    Invoke-SafeCommand -Command {
        Get-ActiveSyncDeviceAccessRule -ErrorAction SilentlyContinue | Select-Object Identity, Characteristic, QueryString,
        AccessLevel, @{N='CollectedDate';E={Get-Date}}
    } -Description "Mobile Device Access Rules" -Category "EXO_MobileDeviceRules"

    # Mobile Device Mailbox Policies
    Invoke-SafeCommand -Command {
        Get-ActiveSyncMailboxPolicy -ErrorAction SilentlyContinue | Select-Object Name, AllowNonProvisionableDevices,
        PasswordEnabled, AlphanumericPasswordRequired, PasswordRecoveryEnabled, DeviceEncryptionEnabled,
        AttachmentsEnabled, MaxAttachmentSize, AllowStorageCard, AllowCamera, AllowWiFi, AllowBluetooth,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Mobile Device Mailbox Policies" -Category "EXO_MobileDevicePolicies"

    # OWA Policies
    Invoke-SafeCommand -Command {
        Get-OwaMailboxPolicy -ErrorAction SilentlyContinue | Select-Object Name, DirectFileAccessOnPublicComputersEnabled,
        DirectFileAccessOnPrivateComputersEnabled, WebReadyDocumentViewingOnPublicComputersEnabled,
        ForceWebReadyDocumentViewingFirstOnPublicComputers, ActiveSyncIntegrationEnabled,
        AllowOfflineOn, ExternalImageProxyEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "OWA Policies" -Category "EXO_OWAPolicies"

    # Retention Policies
    Invoke-SafeCommand -Command {
        Get-RetentionPolicy -ErrorAction SilentlyContinue | Select-Object Name, RetentionPolicyTagLinks, IsDefault,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Retention Policies" -Category "EXO_RetentionPolicies"

    # Data Loss Prevention Policies
    Invoke-SafeCommand -Command {
        Get-DlpPolicy -ErrorAction SilentlyContinue | Select-Object Name, State, Mode, Description,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "DLP Policies" -Category "EXO_DLPPolicies"

    # Quarantine Policies
    Invoke-SafeCommand -Command {
        Get-QuarantinePolicy -ErrorAction SilentlyContinue | Select-Object Name,
        EndUserQuarantinePermissionsValue, ESNEnabled, @{N='CollectedDate';E={Get-Date}}
    } -Description "Quarantine Policies" -Category "EXO_QuarantinePolicies"

    # Address Lists
    Invoke-SafeCommand -Command {
        Get-AddressList -ErrorAction SilentlyContinue | Select-Object Name, RecipientFilter, DisplayName,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Address Lists" -Category "EXO_AddressLists"

    # Global Address Lists
    Invoke-SafeCommand -Command {
        Get-GlobalAddressList -ErrorAction SilentlyContinue | Select-Object Name, RecipientFilter,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Global Address Lists" -Category "EXO_GlobalAddressLists"

    # Offline Address Books
    Invoke-SafeCommand -Command {
        Get-OfflineAddressBook -ErrorAction SilentlyContinue | Select-Object Name, AddressLists, IsDefault,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Offline Address Books" -Category "EXO_OfflineAddressBooks"

    # Organization Relationships (Federation)
    Invoke-SafeCommand -Command {
        Get-OrganizationRelationship -ErrorAction SilentlyContinue | Select-Object Name, DomainNames,
        FreeBusyAccessEnabled, FreeBusyAccessLevel, FreeBusyAccessScope, MailboxMoveEnabled,
        DeliveryReportEnabled, MailTipsAccessEnabled, MailTipsAccessLevel, MailTipsAccessScope,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Organization Relationships" -Category "EXO_OrganizationRelationships"

    # Sharing Policies
    Invoke-SafeCommand -Command {
        Get-SharingPolicy -ErrorAction SilentlyContinue | Select-Object Name, Domains, Enabled, Default,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Sharing Policies" -Category "EXO_SharingPolicies"

    # Role Assignment Policies
    Invoke-SafeCommand -Command {
        Get-RoleAssignmentPolicy -ErrorAction SilentlyContinue | Select-Object Name, Description, IsDefault, AssignedRoles,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Role Assignment Policies" -Category "EXO_RoleAssignmentPolicies"

    # Audit Configuration
    Invoke-SafeCommand -Command {
        Get-AdminAuditLogConfig -ErrorAction SilentlyContinue | Select-Object AdminAuditLogEnabled, AdminAuditLogCmdlets,
        AdminAuditLogParameters, AdminAuditLogExcludedCmdlets, LogLevel,
        @{N='CollectedDate';E={Get-Date}}
    } -Description "Admin Audit Log Configuration" -Category "EXO_AdminAuditConfig"

    # Microsoft Graph data (if connected)
    if ($Script:ConnectedToGraph) {
        Invoke-SafeCommand -Command {
            Import-Module Microsoft.Graph.Identity.DirectoryManagement -Force -ErrorAction Stop
            $org = Get-MgOrganization -ErrorAction Stop
            $orgInfo = @{
                DisplayName = $org.DisplayName
                TenantType = $org.TenantType
                CountryLetterCode = $org.CountryLetterCode
                CreatedDateTime = $org.CreatedDateTime
                TechnicalNotificationMails = $org.TechnicalNotificationMails -join ", "
                CollectedDate = Get-Date
            }
            return [PSCustomObject]$orgInfo
        } -Description "Azure AD Organization Info" -Category "AAD_OrganizationInfo"
    }

    Write-LogProgress "Completed comprehensive Exchange Online data collection"
}

# Function to export data to CSV - COMPLETELY REWRITTEN
function Export-ToCSV {
    Write-LogProgress "Exporting data to CSV format (creating separate files per category)"

    $csvFileCount = 0
    $csvOutputFolder = Join-Path $OutputPath "CSV_Export_$Script:Timestamp"

    # Create CSV output folder
    if (-not (Test-Path $csvOutputFolder)) {
        New-Item -ItemType Directory -Path $csvOutputFolder -Force | Out-Null
    }

    foreach ($category in $Script:ReportData.Keys | Sort-Object) {
        $data = $Script:ReportData[$category]

        if ($null -eq $data) {
            Write-Verbose "Skipping null category: $category"
            continue
        }

        # Skip empty arrays
        if ($data -is [Array] -and $data.Count -eq 0) {
            Write-Verbose "Skipping empty category: $category"
            continue
        }

        $csvPath = Join-Path $csvOutputFolder "${category}.csv"

        try {
            if ($data -is [Array] -and $data.Count -gt 0) {
                # Export array of objects
                $data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                Write-Verbose "Exported $($data.Count) items to $csvPath"
                $csvFileCount++
            } elseif ($data -is [PSCustomObject] -or $data -is [System.Management.Automation.PSCustomObject]) {
                # Export single object
                $data | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                Write-Verbose "Exported single object to $csvPath"
                $csvFileCount++
            } else {
                # Handle other types by converting to string representation
                $outputData = [PSCustomObject]@{
                    Category = $category
                    Value = $data.ToString()
                    Type = $data.GetType().FullName
                    CollectedDate = Get-Date
                }
                $outputData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
                Write-Verbose "Exported non-standard data type to $csvPath"
                $csvFileCount++
            }
        }
        catch {
            Write-Warning "Failed to export $category to CSV: $($_.Exception.Message)"
        }
    }

    Write-LogProgress "CSV reports saved to: $csvOutputFolder ($csvFileCount files created)"
    return $csvOutputFolder
}

# Function to generate HTML report
function Export-ToHTML {
    Write-LogProgress "Generating comprehensive HTML report"

    # Load System.Web assembly for HTML encoding at the beginning
    Add-Type -AssemblyName System.Web

    # Determine environment type for report title
    $envType = switch ($Environment) {
        "OnPremises" { "On-Premises Exchange" }
        "Online" { "Exchange Online" }
        "Both" { "Hybrid Exchange Environment" }
    }

    $htmlContent = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>$envType Comprehensive Infrastructure Documentation</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f5f5; }
        .container { max-width: 1400px; margin: 0 auto; background-color: white; padding: 30px; border-radius: 10px; box-shadow: 0 0 20px rgba(0,0,0,0.1); }
        h1 { color: #2c3e50; text-align: center; margin-bottom: 30px; border-bottom: 3px solid #3498db; padding-bottom: 10px; }
        h2 { color: #34495e; margin-top: 30px; margin-bottom: 15px; padding: 10px; background-color: #ecf0f1; border-left: 5px solid #3498db; }
        h3 { color: #2c3e50; margin-top: 20px; margin-bottom: 10px; }
        .info-box { background-color: #e8f4fd; border: 1px solid #bee5eb; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .warning-box { background-color: #fff3cd; border: 1px solid #ffeaa7; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .success-box { background-color: #d4edda; border: 1px solid #c3e6cb; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .critical-box { background-color: #f8d7da; border: 1px solid #f5c6cb; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .online-box { background-color: #e1f5fe; border: 1px solid #81d4fa; border-radius: 5px; padding: 15px; margin: 10px 0; }
        .onprem-box { background-color: #f3e5f5; border: 1px solid #ce93d8; border-radius: 5px; padding: 15px; margin: 10px 0; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; background-color: white; font-size: 0.9em; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: left; word-wrap: break-word; }
        th { background-color: #3498db; color: white; font-weight: bold; position: sticky; top: 0; }
        tr:nth-child(even) { background-color: #f8f9fa; }
        tr:hover { background-color: #e8f4fd; }
        .timestamp { color: #7f8c8d; font-style: italic; text-align: center; margin-top: 30px; }
        .summary-stats { display: flex; justify-content: space-around; margin: 20px 0; flex-wrap: wrap; }
        .stat-box { background-color: #3498db; color: white; padding: 20px; border-radius: 10px; text-align: center; min-width: 150px; margin: 5px; }
        .stat-box.online { background-color: #2196f3; }
        .stat-box.onprem { background-color: #9c27b0; }
        .stat-box.critical { background-color: #e74c3c; }
        .stat-box.warning { background-color: #f39c12; }
        .stat-number { font-size: 2em; font-weight: bold; }
        .stat-label { font-size: 0.9em; margin-top: 5px; }
        .no-data { color: #7f8c8d; font-style: italic; text-align: center; padding: 20px; }
        .collapsible { background-color: #3498db; color: white; cursor: pointer; padding: 15px; width: 100%; border: none; text-align: left; outline: none; font-size: 16px; margin-top: 10px; border-radius: 5px; }
        .collapsible:hover { background-color: #2980b9; }
        .collapsible.online { background-color: #2196f3; }
        .collapsible.online:hover { background-color: #1976d2; }
        .collapsible.onprem { background-color: #9c27b0; }
        .collapsible.onprem:hover { background-color: #7b1fa2; }
        .collapsible.critical { background-color: #e74c3c; }
        .collapsible.critical:hover { background-color: #c0392b; }
        .content { padding: 0; display: none; overflow: hidden; background-color: #f8f9fa; border-radius: 0 0 5px 5px; }
        .content.show { display: block; padding: 15px; }
        .environment-badge { display: inline-block; padding: 5px 10px; border-radius: 15px; font-size: 0.8em; font-weight: bold; margin-left: 10px; }
        .badge-online { background-color: #2196f3; color: white; }
        .badge-onprem { background-color: #9c27b0; color: white; }
        .badge-critical { background-color: #e74c3c; color: white; }
        .cert-expired { background-color: #ffebee !important; color: #c62828; font-weight: bold; }
        .cert-expiring { background-color: #fff3e0 !important; color: #ef6c00; }
        .cert-valid { background-color: #e8f5e9 !important; color: #2e7d32; }
        .table-container { overflow-x: auto; max-height: 600px; overflow-y: auto; }
        .expand-all-btn { background-color: #27ae60; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; margin: 10px 5px; }
        .collapse-all-btn { background-color: #e67e22; color: white; padding: 10px 20px; border: none; border-radius: 5px; cursor: pointer; margin: 10px 5px; }
        .expand-all-btn:hover { background-color: #229954; }
        .collapse-all-btn:hover { background-color: #ca6f1e; }
    </style>
    <script>
        function toggleContent(element) {
            var content = element.nextElementSibling;
            content.classList.toggle('show');
            element.textContent = content.classList.contains('show') ?
                element.textContent.replace('>', 'v') :
                element.textContent.replace('v', '>');
        }

        function expandAll() {
            var contents = document.querySelectorAll('.content');
            var buttons = document.querySelectorAll('.collapsible');
            contents.forEach(function(content) {
                content.classList.add('show');
            });
            buttons.forEach(function(button) {
                button.textContent = button.textContent.replace('>', 'v');
            });
        }

        function collapseAll() {
            var contents = document.querySelectorAll('.content');
            var buttons = document.querySelectorAll('.collapsible');
            contents.forEach(function(content) {
                content.classList.remove('show');
            });
            buttons.forEach(function(button) {
                button.textContent = button.textContent.replace('v', '>');
            });
        }
    </script>
</head>
<body>
    <div class="container">
        <h1>$envType Comprehensive Infrastructure Documentation</h1>
        <div class="info-box">
            <strong>Report Generated:</strong> $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')<br>
            <strong>Environment:</strong> $Environment<br>
            <strong>Total Categories:</strong> $($Script:ReportData.Keys.Count)<br>
            <strong>Exchange Online Connected:</strong> $Script:ConnectedToEXO<br>
            <strong>Microsoft Graph Connected:</strong> $Script:ConnectedToGraph<br>
            <strong>Detailed Statistics Included:</strong> $IncludeDetailedStats<br>
            <strong>HTTPS Used (On-Premises):</strong> $UseHTTPS
        </div>
"@

    # Add critical alerts section with improved logic
    $criticalAlerts = @()

    # Check for expired certificates
    if ($Script:ReportData.ContainsKey("ExchangeCertificates")) {
        $certs = $Script:ReportData["ExchangeCertificates"]
        if ($certs -and $certs.Count -gt 0) {
            $expiredCerts = $certs | Where-Object {$_.IsExpired -eq $true}
            $expiringSoonCerts = $certs | Where-Object {
                $_.DaysUntilExpiry -ne $null -and
                $_.DaysUntilExpiry -le 30 -and
                $_.DaysUntilExpiry -gt 0 -and
                $_.IsExpired -ne $true
            }

            if ($expiredCerts -and $expiredCerts.Count -gt 0) {
                $criticalAlerts += "[ALERT] CRITICAL: $($expiredCerts.Count) expired certificate(s) found"
            }
            if ($expiringSoonCerts -and $expiringSoonCerts.Count -gt 0) {
                $criticalAlerts += "[WARNING] WARNING: $($expiringSoonCerts.Count) certificate(s) expiring within 30 days"
            }
        }
    }

    # Check for errors during collection
    if ($Script:ErrorLog.Count -gt 0) {
        $criticalAlerts += "[WARNING] $($Script:ErrorLog.Count) error(s) occurred during data collection"
    }

    if ($criticalAlerts.Count -gt 0) {
        $htmlContent += "<div class='critical-box'><h3>[ALERT] Critical Alerts</h3><ul>"
        foreach ($alert in $criticalAlerts) {
            $htmlContent += "<li>$alert</li>"
        }
        $htmlContent += "</ul></div>"
    } else {
        $htmlContent += "<div class='success-box'><h3>[OK] No Critical Issues Detected</h3><p>All checks passed successfully.</p></div>"
    }

    # Add summary statistics if available
    $hasStats = $Script:ReportData.ContainsKey("MailboxStatistics") -or $Script:ReportData.ContainsKey("EXO_MailboxStatistics")
    if ($hasStats) {
        $htmlContent += "<h2>[STATS] Summary Statistics</h2><div class='summary-stats'>"

        if ($Script:ReportData.ContainsKey("MailboxStatistics")) {
            $stats = $Script:ReportData["MailboxStatistics"]
            $htmlContent += "<div class='stat-box onprem'><div class='stat-number'>$($stats.TotalMailboxes)</div><div class='stat-label'>Total Mailboxes (On-Prem)</div></div>"
            $htmlContent += "<div class='stat-box onprem'><div class='stat-number'>$($stats.UserMailboxes)</div><div class='stat-label'>User Mailboxes</div></div>"
            $htmlContent += "<div class='stat-box onprem'><div class='stat-number'>$($stats.SharedMailboxes)</div><div class='stat-label'>Shared Mailboxes</div></div>"
        }

        if ($Script:ReportData.ContainsKey("EXO_MailboxStatistics")) {
            $stats = $Script:ReportData["EXO_MailboxStatistics"]
            $htmlContent += "<div class='stat-box online'><div class='stat-number'>$($stats.TotalMailboxes)</div><div class='stat-label'>Total Mailboxes (Online)</div></div>"
            $htmlContent += "<div class='stat-box online'><div class='stat-number'>$($stats.UserMailboxes)</div><div class='stat-label'>User Mailboxes</div></div>"
            $htmlContent += "<div class='stat-box online'><div class='stat-number'>$($stats.SharedMailboxes)</div><div class='stat-label'>Shared Mailboxes</div></div>"
        }

        if ($Script:ReportData.ContainsKey("ExchangeServers")) {
            $serverCount = $Script:ReportData["ExchangeServers"].Count
            $htmlContent += "<div class='stat-box onprem'><div class='stat-number'>$serverCount</div><div class='stat-label'>Exchange Servers</div></div>"
        }

        if ($Script:ReportData.ContainsKey("ExchangeCertificates")) {
            $certs = $Script:ReportData["ExchangeCertificates"]
            if ($certs -and $certs.Count -gt 0) {
                $certCount = $certs.Count
                $expiredCount = ($certs | Where-Object {$_.IsExpired -eq $true}).Count
                if ($expiredCount -gt 0) {
                    $htmlContent += "<div class='stat-box critical'><div class='stat-number'>$expiredCount</div><div class='stat-label'>Expired Certificates</div></div>"
                }
                $htmlContent += "<div class='stat-box onprem'><div class='stat-number'>$certCount</div><div class='stat-label'>Total Certificates</div></div>"
            }
        }

        if ($Script:ReportData.ContainsKey("EXO_AcceptedDomains")) {
            $domainCount = $Script:ReportData["EXO_AcceptedDomains"].Count
            $htmlContent += "<div class='stat-box online'><div class='stat-number'>$domainCount</div><div class='stat-label'>Accepted Domains (Online)</div></div>"
        }

        $htmlContent += "</div>"
    }

    # Add expand/collapse all buttons
    $htmlContent += @"
    <div style="text-align: center; margin: 20px 0;">
        <button class="expand-all-btn" onclick="expandAll()">v Expand All Sections</button>
        <button class="collapse-all-btn" onclick="collapseAll()">> Collapse All Sections</button>
    </div>
"@

    # Generate sections for each category with enhanced styling
    foreach ($category in $Script:ReportData.Keys | Sort-Object) {
        $data = $Script:ReportData[$category]
        $displayName = $category -replace "_", " " -replace "EXO", "Exchange Online"

        # Determine environment type and criticality for styling
        $envClass = "onprem"
        $envBadge = "<span class='environment-badge badge-onprem'>On-Premises</span>"

        if ($category.StartsWith("EXO_") -or $category.StartsWith("AAD_")) {
            $envClass = "online"
            $envBadge = "<span class='environment-badge badge-online'>Exchange Online</span>"
        }

        # Mark critical categories
        $isCritical = $false
        if ($category -eq "ExchangeCertificates" -and $criticalAlerts.Count -gt 0) {
            $envClass = "critical"
            $envBadge += "<span class='environment-badge badge-critical'>[WARNING] Critical</span>"
            $isCritical = $true
        }

        $htmlContent += "<button class='collapsible $envClass' onclick='toggleContent(this)'>> $displayName $envBadge</button>"
        $htmlContent += "<div class='content'>"

        if ($data -and (($data -is [Array] -and $data.Count -gt 0) -or ($data -isnot [Array]))) {
            if ($data -is [Array] -and $data[0] -is [PSCustomObject]) {
                # Create table for structured data
                $properties = $data[0].PSObject.Properties.Name
                $htmlContent += "<div class='table-container'><table><thead><tr>"
                foreach ($prop in $properties) {
                    $htmlContent += "<th>$([System.Web.HttpUtility]::HtmlEncode($prop))</th>"
                }
                $htmlContent += "</tr></thead><tbody>"

                foreach ($item in $data) {
                    $rowClass = ""

                    # Special formatting for certificates with improved logic
                    if ($category -eq "ExchangeCertificates") {
                        if ($item.IsExpired -eq $true) {
                            $rowClass = "cert-expired"
                        } elseif ($item.DaysUntilExpiry -ne $null -and $item.DaysUntilExpiry -le 30 -and $item.DaysUntilExpiry -gt 0) {
                            $rowClass = "cert-expiring"
                        } elseif ($item.DaysUntilExpiry -ne $null -and $item.DaysUntilExpiry -gt 30) {
                            $rowClass = "cert-valid"
                        }
                    }

                    $htmlContent += "<tr class='$rowClass'>"
                    foreach ($prop in $properties) {
                        $value = $item.$prop
                        if ($null -eq $value) {
                            $value = ""
                        } elseif ($value -is [Array]) {
                            $value = $value -join ", "
                        }
                        $htmlContent += "<td>$([System.Web.HttpUtility]::HtmlEncode($value))</td>"
                    }
                    $htmlContent += "</tr>"
                }
                $htmlContent += "</tbody></table></div>"
            } elseif ($data -is [PSCustomObject]) {
                # Single object - display as key-value pairs
                $htmlContent += "<table><thead><tr><th>Property</th><th>Value</th></tr></thead><tbody>"
                foreach ($prop in $data.PSObject.Properties) {
                    $value = $prop.Value
                    if ($null -eq $value) {
                        $value = ""
                    } elseif ($value -is [Array]) {
                        $value = $value -join ", "
                    }
                    $htmlContent += "<tr><td><strong>$([System.Web.HttpUtility]::HtmlEncode($prop.Name))</strong></td><td>$([System.Web.HttpUtility]::HtmlEncode($value))</td></tr>"
                }
                $htmlContent += "</tbody></table>"
            } else {
                $htmlContent += "<pre>$([System.Web.HttpUtility]::HtmlEncode(($data | Out-String)))</pre>"
            }
        } else {
            $htmlContent += "<div class='no-data'>No data available for this category</div>"
        }

        $htmlContent += "</div>"
    }

    # Add error log section if there were errors
    if ($Script:ErrorLog.Count -gt 0) {
        $htmlContent += "<h2>[WARNING] Error Log</h2>"
        $htmlContent += "<div class='warning-box'>"
        $htmlContent += "<p>The following errors occurred during data collection:</p>"
        $htmlContent += "<table><thead><tr><th>Timestamp</th><th>Category</th><th>Description</th><th>Error Message</th></tr></thead><tbody>"
        foreach ($error in $Script:ErrorLog) {
            $htmlContent += "<tr>"
            $htmlContent += "<td>$($error.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))</td>"
            $htmlContent += "<td>$([System.Web.HttpUtility]::HtmlEncode($error.Category))</td>"
            $htmlContent += "<td>$([System.Web.HttpUtility]::HtmlEncode($error.Description))</td>"
            $htmlContent += "<td>$([System.Web.HttpUtility]::HtmlEncode($error.ErrorMessage))</td>"
            $htmlContent += "</tr>"
        }
        $htmlContent += "</tbody></table></div>"
    }

    $htmlContent += @"
        <div class="timestamp">
            Comprehensive report generated on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') by Exchange Documentation Script v3.1 (FIXED)<br>
            This report includes SMTP relays, EWS certificates, and all critical Exchange configurations<br>
            Script fixes applied: EXO connection, CSV export, public folders, performance optimization, error handling
        </div>
    </div>
</body>
</html>
"@

    try {
        $htmlContent | Out-File -FilePath $Script:HTMLPath -Encoding UTF8 -Force
        Write-LogProgress "Comprehensive HTML report saved to: $Script:HTMLPath"
        return $true
    }
    catch {
        Write-Error "Failed to save HTML report: $($_.Exception.Message)"
        return $false
    }
}

# Main execution function
function Start-ExchangeDocumentation {
    Write-LogProgress "Starting Comprehensive Exchange Infrastructure Documentation (FIXED VERSION)" "Initializing"

    # Check required modules
    if (-not (Test-RequiredModules -Environment $Environment)) {
        return
    }

    # Create output directory if it doesn't exist
    if (-not (Test-Path $OutputPath)) {
        try {
            New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
            Write-Verbose "Created output directory: $OutputPath"
        }
        catch {
            Write-Error "Failed to create output directory: $($_.Exception.Message)"
            return
        }
    }

    $onPremSession = $null

    try {
        # Connect and collect data based on environment
        switch ($Environment) {
            "OnPremises" {
                if (-not $ExchangeServer -or [string]::IsNullOrWhiteSpace($ExchangeServer)) {
                    do {
                        $ExchangeServer = Read-Host "Enter Exchange Server FQDN (e.g., mail.contoso.com)"
                        if ([string]::IsNullOrWhiteSpace($ExchangeServer)) {
                            Write-Warning "Server name cannot be empty"
                        }
                    } while ([string]::IsNullOrWhiteSpace($ExchangeServer))
                }

                $onPremSession = Connect-ExchangeOnPremises -Server $ExchangeServer -Cred $Credential -UseHTTPS $UseHTTPS
                if ($onPremSession) {
                    Get-ExchangeOnPremisesData
                } else {
                    Write-Error "Failed to connect to Exchange On-Premises. Cannot continue."
                    return
                }
            }
            "Online" {
                if (Connect-ExchangeOnline -TenantId $TenantId -AppId $AppId -CertThumbprint $CertificateThumbprint) {
                    # Also try to connect to Microsoft Graph for additional data
                    Connect-MicrosoftGraph -TenantId $TenantId | Out-Null
                    Get-ExchangeOnlineData
                } else {
                    Write-Error "Failed to connect to Exchange Online. Cannot continue."
                    return
                }
            }
            "Both" {
                # On-Premises first
                if (-not $ExchangeServer -or [string]::IsNullOrWhiteSpace($ExchangeServer)) {
                    $ExchangeServer = Read-Host "Enter Exchange Server FQDN (or press Enter to skip On-Premises)"
                }

                if (-not [string]::IsNullOrWhiteSpace($ExchangeServer)) {
                    $onPremSession = Connect-ExchangeOnPremises -Server $ExchangeServer -Cred $Credential -UseHTTPS $UseHTTPS
                    if ($onPremSession) {
                        Get-ExchangeOnPremisesData
                    } else {
                        Write-Warning "Failed to connect to Exchange On-Premises. Continuing with Exchange Online only."
                    }
                }

                # Then Exchange Online
                if (Connect-ExchangeOnline -TenantId $TenantId -AppId $AppId -CertThumbprint $CertificateThumbprint) {
                    Connect-MicrosoftGraph -TenantId $TenantId | Out-Null
                    Get-ExchangeOnlineData
                } else {
                    Write-Warning "Failed to connect to Exchange Online."
                    if (-not $onPremSession) {
                        Write-Error "Failed to connect to both environments. Cannot continue."
                        return
                    }
                }
            }
        }

        # Generate reports
        if ($Script:ReportData.Keys.Count -gt 0) {
            $csvFolder = Export-ToCSV
            $htmlSuccess = Export-ToHTML

            Write-Host "`n" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "COMPREHENSIVE EXCHANGE DOCUMENTATION COMPLETED" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "Environment: $Environment" -ForegroundColor Cyan
            Write-Host "CSV Reports Folder: $csvFolder" -ForegroundColor Yellow
            Write-Host "HTML Report: $Script:HTMLPath" -ForegroundColor Yellow
            Write-Host "Total Categories Documented: $($Script:ReportData.Keys.Count)" -ForegroundColor Cyan
            Write-Host "Exchange Online Connected: $Script:ConnectedToEXO" -ForegroundColor Cyan
            Write-Host "Microsoft Graph Connected: $Script:ConnectedToGraph" -ForegroundColor Cyan
            Write-Host "Detailed Statistics: $IncludeDetailedStats" -ForegroundColor Cyan
            Write-Host "HTTPS Used: $UseHTTPS" -ForegroundColor Cyan
            Write-Host "Errors Encountered: $($Script:ErrorLog.Count)" -ForegroundColor $(if ($Script:ErrorLog.Count -gt 0) { "Yellow" } else { "Green" })
            Write-Host "========================================" -ForegroundColor Green

            # Show critical alerts in console
            if ($Script:ReportData.ContainsKey("ExchangeCertificates")) {
                $certs = $Script:ReportData["ExchangeCertificates"]
                if ($certs -and $certs.Count -gt 0) {
                    $expiredCerts = $certs | Where-Object {$_.IsExpired -eq $true}
                    $expiringSoonCerts = $certs | Where-Object {
                        $_.DaysUntilExpiry -ne $null -and
                        $_.DaysUntilExpiry -le 30 -and
                        $_.DaysUntilExpiry -gt 0 -and
                        $_.IsExpired -ne $true
                    }

                    if ($expiredCerts -and $expiredCerts.Count -gt 0) {
                        Write-Host "[ALERT] CRITICAL: $($expiredCerts.Count) expired certificate(s) found!" -ForegroundColor Red
                        Write-Host "   Review the 'ExchangeCertificates' section in the HTML report" -ForegroundColor Red
                    }
                    if ($expiringSoonCerts -and $expiringSoonCerts.Count -gt 0) {
                        Write-Host "[WARNING]  WARNING: $($expiringSoonCerts.Count) certificate(s) expiring within 30 days!" -ForegroundColor Yellow
                        Write-Host "   Review the 'ExchangeCertificates' section in the HTML report" -ForegroundColor Yellow
                    }
                }
            }

            if ($Script:ErrorLog.Count -gt 0) {
                Write-Host "`n[WARNING]  $($Script:ErrorLog.Count) error(s) occurred during data collection." -ForegroundColor Yellow
                Write-Host "   Review the 'Error Log' section in the HTML report for details" -ForegroundColor Yellow
            }

        } else {
            Write-Warning "No data was collected. Please check your connections and permissions."
            Write-Warning "Review any error messages above."
        }
    }
    catch {
        Write-Error "An error occurred during documentation: $($_.Exception.Message)"
        Write-Error "Stack Trace: $($_.Exception.StackTrace)"
    }
    finally {
        # Cleanup connections - FIXED ORDER: Disconnect cloud services first, then remove sessions
        Write-LogProgress "Cleaning up connections" "Finalizing"

        # Disconnect Exchange Online first
        if ($Script:ConnectedToEXO) {
            try {
                Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
                Write-Verbose "Successfully disconnected from Exchange Online"
            } catch {
                Write-Verbose "Could not disconnect from Exchange Online: $($_.Exception.Message)"
            }
        }

        # Disconnect Microsoft Graph
        if ($Script:ConnectedToGraph) {
            try {
                Disconnect-MgGraph -ErrorAction Stop
                Write-Verbose "Successfully disconnected from Microsoft Graph"
            } catch {
                Write-Verbose "Could not disconnect from Microsoft Graph: $($_.Exception.Message)"
            }
        }

        # Remove on-premises session last
        if ($onPremSession) {
            try {
                Remove-PSSession $onPremSession -ErrorAction Stop
                Write-Verbose "Successfully removed on-premises PowerShell session"
            } catch {
                Write-Verbose "Could not remove on-premises session: $($_.Exception.Message)"
            }
        }

        Write-Progress -Activity "Exchange Comprehensive Documentation" -Completed
        Write-LogProgress "Documentation process completed" "Done" -NoProgress
    }
}

# Execute the main function
Start-ExchangeDocumentation
