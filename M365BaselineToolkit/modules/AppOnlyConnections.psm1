#Requires -Version 7.0
<#
    Certificate-based, app-only (unattended) connection function for Microsoft
    Graph, Exchange Online, Teams, and SharePoint Online - an alternative to
    the interactive sign-in Connect-BaselineWorkload provides in
    BaselineCore.psm1.

    This module does not modify, replace, or depend on any change to
    BaselineCore.psm1 or Connect-BaselineWorkload - it is a parallel path.
    Only Invoke-M365Baseline.AppOnly.ps1 calls into this module; the
    interactive Invoke-M365Baseline.ps1 never imports it and is unaffected by
    its existence.

    Design choice per the app-only work's spec: whichever certificate input
    form is given (thumbprint from a local store, a .pfx file, or a
    caller-supplied X509Certificate2 object - e.g. from Key Vault) is resolved
    into exactly ONE X509Certificate2 object, and that same object instance is
    passed to every Connect-* call. This avoids a documented SharePoint Online
    Management Shell quirk where -CertificateThumbprint always looks in the
    LocalMachine store regardless of where the certificate actually lives -
    passing the resolved object instead of a thumbprint string sidesteps that
    entirely for all four services, not just SharePoint.

    See APP-ONLY-AUTH section in README.md for one-time tenant setup
    (app registration, certificate, API permissions, directory role
    assignments) before this module can be used against a real tenant.
#>

Set-StrictMode -Version Latest

function Resolve-M365BaselineAppOnlyCertificate {
    <#
    .SYNOPSIS
        Internal: resolves any of the three supported certificate input forms
        into a single validated X509Certificate2 object.
    #>
    [CmdletBinding(DefaultParameterSetName = 'Thumbprint')]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param(
        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')]
        [string]$CertificateThumbprint,

        [Parameter(ParameterSetName = 'Thumbprint')]
        [ValidateSet('CurrentUser', 'LocalMachine')]
        [string]$CertificateStoreLocation = 'CurrentUser',

        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]$CertificatePath,

        [Parameter(ParameterSetName = 'File')]
        [securestring]$CertificatePassword,

        [Parameter(Mandatory, ParameterSetName = 'Object')]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    switch ($PSCmdlet.ParameterSetName) {
        'Object' {
            if (-not $Certificate.HasPrivateKey) {
                throw "The supplied -Certificate object has no private key. App-only (client-credentials) auth requires the private key, not just the public certificate."
            }
            return $Certificate
        }
        'File' {
            if (-not (Test-Path -LiteralPath $CertificatePath)) {
                throw "Certificate file not found: $CertificatePath"
            }
            $cert = if ($CertificatePassword) {
                [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $CertificatePassword)
            }
            else {
                [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath)
            }
            if (-not $cert.HasPrivateKey) {
                throw "Certificate file '$CertificatePath' does not contain a private key. Export the .pfx with 'include private key', not just the public .cer."
            }
            return $cert
        }
        'Thumbprint' {
            $storePath = "Cert:\$CertificateStoreLocation\My"
            $cert = Get-ChildItem -Path $storePath -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $CertificateThumbprint }
            if (-not $cert) {
                throw "No certificate with thumbprint '$CertificateThumbprint' was found in $storePath. Import it there, pass -CertificateStoreLocation LocalMachine if it's in the machine store instead of the user store, or use -CertificatePath/-Certificate instead."
            }
            if (-not $cert.HasPrivateKey) {
                throw "Certificate '$CertificateThumbprint' in $storePath was found but has no private key available. Re-import the .pfx (not just the public .cer)."
            }
            return $cert
        }
    }
}

function Connect-M365BaselineServicesAppOnly {
    <#
    .SYNOPSIS
        Connects to one or more of Graph/ExchangeOnline/Teams/SharePointOnline
        using app-only (certificate) authentication - never an interactive
        sign-in, never a WAM/browser prompt of any kind.
    .DESCRIPTION
        Resolves the given certificate input (thumbprint, .pfx file, or a
        pre-built X509Certificate2 object) into one certificate object, then
        connects only to the services listed in -Services, passing that same
        certificate object to every Connect-* call.

        Connects in the order -Services is given - callers should order it
        themselves first, e.g. via BaselineCore.psm1's exported
        Get-BaselineConnectionOrder (Graph before ExchangeOnline, for the same
        documented MSAL version-pinning conflict the interactive script's
        Connect-BaselineWorkload observes). This function does not re-sort,
        to avoid duplicating that ordering logic.

        No interactive fallback exists anywhere in this function or module.
        If a service's app-only connection fails, this throws - the caller
        (Invoke-M365Baseline.AppOnly.ps1) does not retry interactively, since
        that would defeat the point of an unattended script.
    .PARAMETER Services
        One or more of 'Graph', 'ExchangeOnline', 'Teams', 'SharePointOnline'.
    .PARAMETER AppId
        The Entra app registration's Application (client) ID.
    .PARAMETER TenantId
        Entra tenant id (GUID) or a verified domain (e.g. contoso.onmicrosoft.com).
    .PARAMETER Organization
        Required when 'ExchangeOnline' is in -Services: the tenant's primary
        *.onmicrosoft.com domain (Exchange Online's app-only auth requires
        this specifically, not a GUID).
    .PARAMETER SpoAdminUrl
        Required when 'SharePointOnline' is in -Services: the SharePoint
        admin center URL (e.g. https://contoso-admin.sharepoint.com).
    .PARAMETER CertificateThumbprint
        Thumbprint of a certificate already imported into a local certificate
        store, whose private key matches the public certificate uploaded to
        the app registration.
    .PARAMETER CertificateStoreLocation
        'CurrentUser' (default) or 'LocalMachine' - which store
        -CertificateThumbprint is looked up in.
    .PARAMETER CertificatePath
        Path to a .pfx file containing the certificate and private key, as an
        alternative to the certificate store (e.g. downloaded from Key Vault
        to a temp path at runtime).
    .PARAMETER CertificatePassword
        SecureString password for -CertificatePath. Omit if the .pfx has none.
    .PARAMETER Certificate
        An already-constructed X509Certificate2 object, for a caller that
        manages its own certificate retrieval (Key Vault SDK, etc.).
    .EXAMPLE
        Connect-M365BaselineServicesAppOnly -Services Graph,ExchangeOnline -AppId $appId -TenantId contoso.onmicrosoft.com -Organization contoso.onmicrosoft.com -CertificateThumbprint $thumbprint
    #>
    [CmdletBinding(DefaultParameterSetName = 'Thumbprint')]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'Teams', 'SharePointOnline')]
        [string[]]$Services,

        [Parameter(Mandatory)]
        [string]$AppId,

        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter()]
        [string]$Organization,

        [Parameter()]
        [string]$SpoAdminUrl,

        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')]
        [string]$CertificateThumbprint,

        [Parameter(ParameterSetName = 'Thumbprint')]
        [ValidateSet('CurrentUser', 'LocalMachine')]
        [string]$CertificateStoreLocation = 'CurrentUser',

        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]$CertificatePath,

        [Parameter(ParameterSetName = 'File')]
        [securestring]$CertificatePassword,

        [Parameter(Mandatory, ParameterSetName = 'Object')]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    if (($Services -contains 'ExchangeOnline') -and [string]::IsNullOrWhiteSpace($Organization)) {
        throw "Connecting to ExchangeOnline (app-only) requires -Organization <tenant>.onmicrosoft.com."
    }
    if (($Services -contains 'SharePointOnline') -and [string]::IsNullOrWhiteSpace($SpoAdminUrl)) {
        throw "Connecting to SharePointOnline (app-only) requires -SpoAdminUrl https://<tenant>-admin.sharepoint.com."
    }

    $resolveParams = switch ($PSCmdlet.ParameterSetName) {
        'Thumbprint' {
            @{ CertificateThumbprint = $CertificateThumbprint; CertificateStoreLocation = $CertificateStoreLocation }
        }
        'File' {
            $p = @{ CertificatePath = $CertificatePath }
            if ($CertificatePassword) { $p['CertificatePassword'] = $CertificatePassword }
            $p
        }
        'Object' {
            @{ Certificate = $Certificate }
        }
    }
    $cert = Resolve-M365BaselineAppOnlyCertificate @resolveParams

    foreach ($service in $Services) {
        try {
            switch ($service) {
                'Graph' {
                    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                    Connect-MgGraph -ClientId $AppId -TenantId $TenantId -Certificate $cert -NoWelcome -ErrorAction Stop
                }
                'ExchangeOnline' {
                    Import-Module ExchangeOnlineManagement -ErrorAction Stop
                    Connect-ExchangeOnline -AppId $AppId -Certificate $cert -Organization $Organization -ShowBanner:$false -ErrorAction Stop
                }
                'Teams' {
                    Import-Module MicrosoftTeams -ErrorAction Stop
                    Connect-MicrosoftTeams -ApplicationId $AppId -TenantId $TenantId -Certificate $cert -ErrorAction Stop | Out-Null
                }
                'SharePointOnline' {
                    # Same -UseWindowsPowerShell -Global rationale as the interactive
                    # Connect-BaselineWorkload in BaselineCore.psm1: this module targets
                    # .NET Framework, not PS7's runtime, and its OAuth handling is
                    # unreliable loaded directly into a PS7 process alongside the other
                    # services' MSAL usage.
                    Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -Global -ErrorAction Stop
                    Connect-SPOService -Url $SpoAdminUrl -ApplicationId $AppId -TenantId $TenantId -Certificate $cert -ErrorAction Stop
                }
            }
        }
        catch {
            throw "Failed to connect to $service (app-only, certificate input '$($PSCmdlet.ParameterSetName)'): $($_.Exception.Message). This is one of three distinct problems, each with a different fix - see README.md's 'App-only (certificate) authentication' section, 'Troubleshooting a connection failure' subsection: (1) the certificate has expired, was revoked, or doesn't match what's uploaded to the app registration; (2) the required API permission for $service was not granted or not admin-consented; (3) $service additionally requires a directory role assignment on the app's service principal (Exchange Administrator for ExchangeOnline, Teams Administrator for Teams) that has not been made, separate from API permissions."
        }
        Set-BaselineWorkloadConnectedState -Connection $service -Connected $true
    }
}

Export-ModuleMember -Function @(
    'Connect-M365BaselineServicesAppOnly'
)
