#Requires -Version 7.0
<#
    Non-interactive (app-only / certificate) connection functions for Microsoft
    Graph and Exchange Online, as an alternative to the interactive sign-in
    Connect-BaselineWorkload uses in BaselineCore.psm1.

    Interactive sign-in on Windows goes through MSAL's Web Account Manager (WAM)
    broker, which has proven unreliable on some Entra-joined workstations (see
    README.md's WAM/RuntimeBroker troubleshooting section) - anything from a
    NullReferenceException crash to a native "Sign in with your work account"
    dialog that reports "User canceled authentication" even when nothing was
    deliberately canceled. App-only auth uses the OAuth2 client-credentials
    flow with a certificate instead of a user sign-in, so it never touches WAM,
    never opens a browser or native dialog, and behaves identically on every
    machine, headless or not.

    This does NOT replace Connect-BaselineWorkload - it only covers Graph and
    Exchange Online. Teams and SharePoint Online connections still go through
    the normal interactive path (see APP-ONLY-AUTH.md for why).

    Setup (Entra app registration, API permissions, certificate) is documented
    in APP-ONLY-AUTH.md, one level up from this module.
#>

Set-StrictMode -Version Latest

function Get-BaselineAppOnlyCertificate {
    <#
    .SYNOPSIS
        Internal: resolves and validates a certificate by thumbprint from the
        local certificate store, with a clear error instead of a cryptic MSAL
        failure when it's missing or has no private key.
    #>
    [CmdletBinding()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param(
        [Parameter(Mandatory)]
        [string]$Thumbprint
    )
    $cert = Get-ChildItem -Path 'Cert:\CurrentUser\My' -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $Thumbprint }
    if (-not $cert) {
        $cert = Get-ChildItem -Path 'Cert:\LocalMachine\My' -ErrorAction SilentlyContinue | Where-Object { $_.Thumbprint -eq $Thumbprint }
    }
    if (-not $cert) {
        throw "No certificate with thumbprint '$Thumbprint' was found in Cert:\CurrentUser\My or Cert:\LocalMachine\My. Import the .pfx first (see APP-ONLY-AUTH.md), or use -CertificatePath instead of -CertificateThumbprint."
    }
    if (-not $cert.HasPrivateKey) {
        throw "Certificate '$Thumbprint' was found but has no private key available. You likely imported only the public .cer - re-import the .pfx that contains the private key."
    }
    return $cert
}

function Get-BaselineAppOnlyCertificateFromFile {
    <#
    .SYNOPSIS
        Internal: loads a certificate + private key from a .pfx file.
    #>
    [CmdletBinding()]
    [OutputType([System.Security.Cryptography.X509Certificates.X509Certificate2])]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter()]
        [securestring]$Password
    )
    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Certificate file not found: $Path"
    }
    $cert = if ($Password) {
        [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Path, $Password)
    }
    else {
        [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Path)
    }
    if (-not $cert.HasPrivateKey) {
        throw "Certificate file '$Path' does not contain a private key. Export the .pfx with 'include private key', not just the public .cer."
    }
    return $cert
}

function Connect-BaselineGraphAppOnly {
    <#
    .SYNOPSIS
        Connects to Microsoft Graph using app-only (client credentials + certificate)
        authentication - never an interactive/WAM sign-in.
    .DESCRIPTION
        Requires an Entra app registration with Application (not Delegated) API
        permissions matching the $script:GraphScopes list in BaselineCore.psm1,
        admin-consented, and a certificate whose public key is uploaded to that
        app registration. See APP-ONLY-AUTH.md for the one-time setup steps.
    .PARAMETER TenantId
        Entra tenant id (GUID) or a verified domain (e.g. contoso.onmicrosoft.com).
    .PARAMETER ClientId
        The app registration's Application (client) ID.
    .PARAMETER CertificateThumbprint
        Thumbprint of a certificate already imported into the current user's or
        local machine's certificate store (Cert:\CurrentUser\My or
        Cert:\LocalMachine\My), whose private key matches the public certificate
        uploaded to the app registration.
    .PARAMETER CertificatePath
        Path to a .pfx file containing the certificate and private key, as an
        alternative to installing it in the certificate store first.
    .PARAMETER CertificatePassword
        SecureString password for -CertificatePath. Omit if the .pfx has none.
    .EXAMPLE
        Connect-BaselineGraphAppOnly -TenantId contoso.onmicrosoft.com -ClientId $appId -CertificateThumbprint $thumbprint
    .EXAMPLE
        Connect-BaselineGraphAppOnly -TenantId contoso.onmicrosoft.com -ClientId $appId -CertificatePath ./app-only.pfx -CertificatePassword $securePwd
    #>
    [CmdletBinding(DefaultParameterSetName = 'Thumbprint')]
    param(
        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')]
        [string]$CertificateThumbprint,

        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]$CertificatePath,

        [Parameter(ParameterSetName = 'File')]
        [securestring]$CertificatePassword
    )

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    try {
        if ($PSCmdlet.ParameterSetName -eq 'Thumbprint') {
            [void](Get-BaselineAppOnlyCertificate -Thumbprint $CertificateThumbprint)
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome -ErrorAction Stop
        }
        else {
            $cert = Get-BaselineAppOnlyCertificateFromFile -Path $CertificatePath -Password $CertificatePassword
            Connect-MgGraph -ClientId $ClientId -TenantId $TenantId -Certificate $cert -NoWelcome -ErrorAction Stop
        }
    }
    catch {
        throw "Failed to connect to Graph (app-only): $($_.Exception.Message). Verify the app registration has admin-consented Application permissions matching `$script:GraphScopes in BaselineCore.psm1, and that the certificate's public key is uploaded to the app registration under Certificates & secrets. See APP-ONLY-AUTH.md."
    }
}

function Connect-BaselineExchangeOnlineAppOnly {
    <#
    .SYNOPSIS
        Connects to Exchange Online using app-only (certificate) authentication -
        never an interactive/WAM sign-in.
    .DESCRIPTION
        Requires the same Entra app registration used for Connect-BaselineGraphAppOnly
        to also have the Office 365 Exchange Online API's Exchange.ManageAsApp
        Application permission granted and admin-consented, and its service
        principal assigned an Exchange-capable Entra role (Exchange Administrator
        is the simplest choice covering everything this toolkit's EXO controls
        need). See APP-ONLY-AUTH.md for the one-time setup steps.
    .PARAMETER Organization
        The tenant's primary *.onmicrosoft.com domain (not a GUID - Exchange
        Online's app-only auth requires the verified domain).
    .PARAMETER ClientId
        The app registration's Application (client) ID.
    .PARAMETER CertificateThumbprint
        Same certificate used for Connect-BaselineGraphAppOnly, by thumbprint
        from the local certificate store.
    .PARAMETER CertificatePath
        Path to a .pfx file, as an alternative to the certificate store.
    .PARAMETER CertificatePassword
        SecureString password for -CertificatePath. Omit if the .pfx has none.
    .EXAMPLE
        Connect-BaselineExchangeOnlineAppOnly -Organization contoso.onmicrosoft.com -ClientId $appId -CertificateThumbprint $thumbprint
    #>
    [CmdletBinding(DefaultParameterSetName = 'Thumbprint')]
    param(
        [Parameter(Mandatory)]
        [string]$Organization,

        [Parameter(Mandatory)]
        [string]$ClientId,

        [Parameter(Mandatory, ParameterSetName = 'Thumbprint')]
        [string]$CertificateThumbprint,

        [Parameter(Mandatory, ParameterSetName = 'File')]
        [string]$CertificatePath,

        [Parameter(ParameterSetName = 'File')]
        [securestring]$CertificatePassword
    )

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    try {
        $eopParams = @{
            AppId        = $ClientId
            Organization = $Organization
            ShowBanner   = $false
            ErrorAction  = 'Stop'
        }
        if ($PSCmdlet.ParameterSetName -eq 'Thumbprint') {
            [void](Get-BaselineAppOnlyCertificate -Thumbprint $CertificateThumbprint)
            $eopParams['CertificateThumbprint'] = $CertificateThumbprint
        }
        else {
            # Validated eagerly (private key present, file exists) even though
            # Connect-ExchangeOnline also takes the raw file path itself, so a bad
            # cert path/password fails with our clear message, not EXO's own.
            [void](Get-BaselineAppOnlyCertificateFromFile -Path $CertificatePath -Password $CertificatePassword)
            $eopParams['CertificateFilePath'] = $CertificatePath
            if ($CertificatePassword) { $eopParams['CertificatePassword'] = $CertificatePassword }
        }
        Connect-ExchangeOnline @eopParams
    }
    catch {
        throw "Failed to connect to Exchange Online (app-only): $($_.Exception.Message). Verify the app registration has admin-consented the Exchange.ManageAsApp Application permission and its service principal has an Exchange-capable Entra role (e.g. Exchange Administrator). See APP-ONLY-AUTH.md."
    }
}

Export-ModuleMember -Function @(
    'Connect-BaselineGraphAppOnly',
    'Connect-BaselineExchangeOnlineAppOnly'
)
