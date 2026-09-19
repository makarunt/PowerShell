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
    its existence, even for the SharePoint child-process mechanism below.

    Design choice per the app-only work's spec: whichever certificate input
    form is given (thumbprint from a local store, a .pfx file, or a
    caller-supplied X509Certificate2 object - e.g. from Key Vault) is resolved
    into exactly ONE X509Certificate2 object, used for Graph/ExchangeOnline/
    Teams directly.

    SharePoint is a deliberate exception to "pass the same object to all four
    services," confirmed necessary via real-tenant troubleshooting: Connect-
    SPOService reached through Import-Module -UseWindowsPowerShell (the same
    WinCompat mechanism the interactive script's Connect-BaselineWorkload
    uses) fails OAuth certificate authentication even with a fully correct
    app registration, permissions, role assignment, and a certificate proven
    to work via a plain native Windows PowerShell 5.1 session - while the
    exact same cert/app/permissions succeed immediately in a genuinely
    separate, directly-spawned Windows PowerShell 5.1 process. The most
    likely explanation is that WinCompat's DCOM-activated background host
    process can enumerate the certificate (list/metadata access) but cannot
    use its private key (a different, more restricted permission) under
    whatever identity/token context that background host actually runs
    under - separate from the calling user's own token, unlike a directly
    spawned child process which inherits it exactly. So SharePoint gets its
    own connection path here: a real child powershell.exe process, talking
    to this module over a line-based JSON protocol on its stdin/stdout, with
    Get-SPOTenant/Set-SPOTenant/Get-SPOBrowserIdleSignOut/
    Set-SPOBrowserIdleSignOut/Disconnect-SPOService defined here as proxy
    functions that forward to it - transparent to the unmodified
    SharePointOnlineControls.psm1, which keeps calling those names exactly as
    it always has, unaware anything changed underneath.

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

# ---------------------------------------------------------------------------
# SharePoint child-process bridge
# ---------------------------------------------------------------------------

$script:SpoChildProcess = $null
$script:SpoChildScriptPath = $null
$script:SpoChildPfxPath = $null

# Runs inside a genuinely separate, native Windows PowerShell 5.1 process
# (never WinCompat) - reads one JSON command per line from stdin, executes
# the real SharePoint Online Management Shell cmdlet, writes one JSON
# response per line to stdout. Kept deliberately tiny: it only implements
# the exact operations SharePointOnlineControls.psm1 (unmodified) needs.
$script:SpoChildServerScript = @'
$ErrorActionPreference = "Stop"
# Recompute PSModulePath from the User/Machine registry-level defaults,
# discarding whatever this process inherited (normally the PS7 parent's own
# PSModulePath, which points at .NET Core module binaries incompatible with
# this .NET Framework runtime and breaks auto-loading of core modules like
# Microsoft.PowerShell.Security - confirmed via real-tenant testing).
# Belt-and-suspenders with Start-BaselineSpoChildProcess already removing it
# from this process's environment before launch.
$env:PSModulePath = @(
    [System.Environment]::GetEnvironmentVariable("PSModulePath", "User")
    [System.Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
) -join ";"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Import-Module Microsoft.Online.SharePoint.PowerShell -ErrorAction Stop

function Write-SpoResponse {
    param($Success, $Result, $ErrorMessage)
    $obj = @{ Success = [bool]$Success }
    if ($Success) { $obj.Result = $Result } else { $obj.Error = [string]$ErrorMessage }
    $json = $obj | ConvertTo-Json -Compress -Depth 6
    [Console]::Out.WriteLine($json)
    [Console]::Out.Flush()
}

while ($true) {
    $line = [Console]::In.ReadLine()
    if ($null -eq $line) { break }
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $shouldExit = $false
    try {
        $request = $line | ConvertFrom-Json
        $result = $null
        switch ($request.Command) {
            'Connect' {
                $p = $request.Params
                $securePwd = ConvertTo-SecureString -String $p.CertificatePassword -AsPlainText -Force
                Connect-SPOService -Url $p.Url -ApplicationId $p.ApplicationId -TenantId $p.TenantId `
                    -CertificatePath $p.CertificatePath -CertificatePassword $securePwd -ErrorAction Stop
                $result = @{ connected = $true }
            }
            'GetTenant' {
                $t = Get-SPOTenant -ErrorAction Stop
                $result = @{
                    SharingCapability                 = [string]$t.SharingCapability
                    DefaultSharingLinkType             = [string]$t.DefaultSharingLinkType
                    DefaultLinkPermission              = [string]$t.DefaultLinkPermission
                    RequireAnonymousLinksExpireInDays  = [int]$t.RequireAnonymousLinksExpireInDays
                    LegacyAuthProtocolsEnabled         = [bool]$t.LegacyAuthProtocolsEnabled
                }
            }
            'SetTenant' {
                $setParams = @{ ErrorAction = 'Stop' }
                foreach ($prop in $request.Params.PSObject.Properties) {
                    $setParams[$prop.Name] = $prop.Value
                }
                Set-SPOTenant @setParams
                $result = @{ ok = $true }
            }
            'GetIdleSignOut' {
                $c = Get-SPOBrowserIdleSignOut -ErrorAction Stop
                $result = @{
                    Enabled             = [bool]$c.Enabled
                    WarnAfterMinutes    = [int]([timespan]$c.WarnAfter).TotalMinutes
                    SignOutAfterMinutes = [int]([timespan]$c.SignOutAfter).TotalMinutes
                }
            }
            'SetIdleSignOut' {
                Set-SPOBrowserIdleSignOut -Enabled:([bool]$request.Params.Enabled) `
                    -WarnAfter (New-TimeSpan -Minutes ([int]$request.Params.WarnAfterMinutes)) `
                    -SignOutAfter (New-TimeSpan -Minutes ([int]$request.Params.SignOutAfterMinutes)) `
                    -ErrorAction Stop
                $result = @{ ok = $true }
            }
            'Disconnect' {
                Disconnect-SPOService -ErrorAction SilentlyContinue
                $result = @{ ok = $true }
            }
            'Exit' {
                $result = @{ ok = $true }
                $shouldExit = $true
            }
            default {
                throw "Unknown command: $($request.Command)"
            }
        }
        Write-SpoResponse -Success $true -Result $result
    }
    catch {
        Write-SpoResponse -Success $false -ErrorMessage $_.Exception.Message
    }
    if ($shouldExit) { break }
}
'@

function Start-BaselineSpoChildProcess {
    <#
    .SYNOPSIS
        Internal: starts (idempotently) the native Windows PowerShell 5.1
        child process the SharePoint proxy functions talk to.
    #>
    [CmdletBinding()]
    param()
    if ($script:SpoChildProcess -and -not $script:SpoChildProcess.HasExited) { return }

    $script:SpoChildScriptPath = Join-Path ([System.IO.Path]::GetTempPath()) "M365BaselineSpoChild-$([guid]::NewGuid()).ps1"
    Set-Content -Path $script:SpoChildScriptPath -Value $script:SpoChildServerScript -Encoding utf8 -ErrorAction Stop

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$($script:SpoChildScriptPath)`""
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    $psi.StandardInputEncoding = [System.Text.UTF8Encoding]::new($false)
    $psi.StandardOutputEncoding = [System.Text.UTF8Encoding]::new($false)
    $psi.UseShellExecute = $false
    $psi.CreateNoWindow = $true
    # PowerShell 7's own PSModulePath (inherited by default, since
    # ProcessStartInfo.EnvironmentVariables starts as a copy of this
    # process's environment) breaks Windows PowerShell 5.1's ability to load
    # its own built-in modules - PS7's module directories point to .NET Core
    # binaries that aren't compatible with WinPS 5.1's .NET Framework
    # runtime, and if they're found first, module auto-load fails with
    # "was found... but the module could not be loaded" for core modules
    # like Microsoft.PowerShell.Security (confirmed via real-tenant testing:
    # ConvertTo-SecureString, called from inside this child process, failed
    # exactly this way before this fix). Removing the inherited value lets
    # native powershell.exe compute its own correct default on startup, the
    # same as if it had been launched with no parent process involved at
    # all - belt-and-suspenders with the same reset done again inside
    # $script:SpoChildServerScript itself, in case something upstream of
    # this process (a system-wide policy, etc.) sets it differently again.
    if ($psi.EnvironmentVariables.ContainsKey('PSModulePath')) {
        $psi.EnvironmentVariables.Remove('PSModulePath')
    }

    $script:SpoChildProcess = [System.Diagnostics.Process]::new()
    $script:SpoChildProcess.StartInfo = $psi
    try {
        [void]$script:SpoChildProcess.Start()
    }
    catch {
        throw "Failed to start the native Windows PowerShell 5.1 child process for SharePoint (powershell.exe): $($_.Exception.Message). Windows PowerShell 5.1 must be installed (it ships with Windows by default) and 'powershell.exe' must be on PATH."
    }
}

function Invoke-BaselineSpoChildCommand {
    <#
    .SYNOPSIS
        Internal: sends one JSON command to the SharePoint child process and
        returns its JSON result, throwing on failure or an unexpected exit.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Command,

        [Parameter()]
        [hashtable]$Params = @{}
    )
    if (-not $script:SpoChildProcess -or $script:SpoChildProcess.HasExited) {
        throw "The SharePoint child process is not running. This is a bug if you've already connected to SharePointOnline this run - otherwise, connect first."
    }
    $request = (@{ Command = $Command; Params = $Params } | ConvertTo-Json -Compress -Depth 6)
    $script:SpoChildProcess.StandardInput.WriteLine($request)
    $script:SpoChildProcess.StandardInput.Flush()

    $responseLine = $script:SpoChildProcess.StandardOutput.ReadLine()
    if ($null -eq $responseLine) {
        $stderr = $script:SpoChildProcess.StandardError.ReadToEnd()
        throw "The SharePoint child process closed unexpectedly while handling '$Command'.$(if ($stderr) { " Stderr: $stderr" })"
    }
    $response = $responseLine | ConvertFrom-Json
    if (-not $response.Success) {
        throw [string]$response.Error
    }
    return $response.Result
}

function Stop-BaselineSpoChildProcess {
    <#
    .SYNOPSIS
        Internal: gracefully asks the SharePoint child process to disconnect
        and exit, then force-kills it if it doesn't within a few seconds.
    #>
    [CmdletBinding()]
    param()
    if (-not $script:SpoChildProcess -or $script:SpoChildProcess.HasExited) { return }
    try {
        Invoke-BaselineSpoChildCommand -Command 'Disconnect' | Out-Null
        Invoke-BaselineSpoChildCommand -Command 'Exit' | Out-Null
    }
    catch {
        Write-Verbose "Non-fatal: SharePoint child process disconnect/exit reported: $($_.Exception.Message)"
    }
    if (-not $script:SpoChildProcess.WaitForExit(5000)) {
        try { $script:SpoChildProcess.Kill() } catch { }
    }
    $script:SpoChildProcess = $null
    if ($script:SpoChildScriptPath -and (Test-Path -LiteralPath $script:SpoChildScriptPath)) {
        Remove-Item -LiteralPath $script:SpoChildScriptPath -Force -ErrorAction SilentlyContinue
    }
    if ($script:SpoChildPfxPath -and (Test-Path -LiteralPath $script:SpoChildPfxPath)) {
        Remove-Item -LiteralPath $script:SpoChildPfxPath -Force -ErrorAction SilentlyContinue
    }
}

# Best-effort cleanup if the process exits without an explicit Disconnect
# (an earlier fatal error, Ctrl+C, etc.) so no orphaned child process or
# temp .pfx is left behind.
$null = Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action {
    if ($script:SpoChildProcess -and -not $script:SpoChildProcess.HasExited) {
        try { $script:SpoChildProcess.Kill() } catch { }
    }
    if ($script:SpoChildScriptPath -and (Test-Path -LiteralPath $script:SpoChildScriptPath)) {
        Remove-Item -LiteralPath $script:SpoChildScriptPath -Force -ErrorAction SilentlyContinue
    }
    if ($script:SpoChildPfxPath -and (Test-Path -LiteralPath $script:SpoChildPfxPath)) {
        Remove-Item -LiteralPath $script:SpoChildPfxPath -Force -ErrorAction SilentlyContinue
    }
} -ErrorAction SilentlyContinue

function Connect-BaselineSpoServiceViaChildProcess {
    <#
    .SYNOPSIS
        Internal: starts the SharePoint child process (if needed) and
        connects it, exporting the already-resolved certificate to a
        short-lived temp .pfx (deleted immediately after) rather than ever
        passing the live object or its private key across the process
        boundary.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Url,

        [Parameter(Mandatory)]
        [string]$AppId,

        [Parameter(Mandatory)]
        [string]$TenantId,

        [Parameter(Mandatory)]
        [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )
    Start-BaselineSpoChildProcess

    $pfxPassword = [System.Guid]::NewGuid().ToString('N')
    $script:SpoChildPfxPath = Join-Path ([System.IO.Path]::GetTempPath()) "M365BaselineSpoCert-$([guid]::NewGuid()).pfx"
    try {
        $pfxBytes = $Certificate.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx, $pfxPassword)
        [System.IO.File]::WriteAllBytes($script:SpoChildPfxPath, $pfxBytes)

        Invoke-BaselineSpoChildCommand -Command 'Connect' -Params @{
            Url                   = $Url
            ApplicationId         = $AppId
            TenantId              = $TenantId
            CertificatePath       = $script:SpoChildPfxPath
            CertificatePassword   = $pfxPassword
        } | Out-Null
    }
    finally {
        if (Test-Path -LiteralPath $script:SpoChildPfxPath) {
            Remove-Item -LiteralPath $script:SpoChildPfxPath -Force -ErrorAction SilentlyContinue
        }
        $script:SpoChildPfxPath = $null
    }
}

function Get-SPOTenant {
    <#
    .SYNOPSIS
        Proxy: forwards to the real Get-SPOTenant running in the SharePoint
        child process. Defined here only so it's visible when
        AppOnlyConnections.psm1 is imported -Global; the interactive script
        never imports this module, so it never sees this function and keeps
        using the real cmdlet exactly as before.
    #>
    [CmdletBinding()]
    param()
    $r = Invoke-BaselineSpoChildCommand -Command 'GetTenant'
    return [pscustomobject]@{
        SharingCapability                = $r.SharingCapability
        DefaultSharingLinkType           = $r.DefaultSharingLinkType
        DefaultLinkPermission            = $r.DefaultLinkPermission
        RequireAnonymousLinksExpireInDays = $r.RequireAnonymousLinksExpireInDays
        LegacyAuthProtocolsEnabled       = $r.LegacyAuthProtocolsEnabled
    }
}

function Set-SPOTenant {
    <#
    .SYNOPSIS
        Proxy: forwards to the real Set-SPOTenant running in the SharePoint
        child process. Accepts exactly the parameters
        SharePointOnlineControls.psm1 actually passes, one at a time.
    #>
    [CmdletBinding()]
    param(
        [Parameter()] [string]$SharingCapability,
        [Parameter()] [string]$DefaultSharingLinkType,
        [Parameter()] [string]$DefaultLinkPermission,
        [Parameter()] [int]$RequireAnonymousLinksExpireInDays,
        [Parameter()] [bool]$LegacyAuthProtocolsEnabled
    )
    $params = @{}
    foreach ($key in @('SharingCapability', 'DefaultSharingLinkType', 'DefaultLinkPermission', 'RequireAnonymousLinksExpireInDays')) {
        if ($PSBoundParameters.ContainsKey($key)) { $params[$key] = $PSBoundParameters[$key] }
    }
    if ($PSBoundParameters.ContainsKey('LegacyAuthProtocolsEnabled')) { $params['LegacyAuthProtocolsEnabled'] = $LegacyAuthProtocolsEnabled }
    Invoke-BaselineSpoChildCommand -Command 'SetTenant' -Params $params | Out-Null
}

function Get-SPOBrowserIdleSignOut {
    <#
    .SYNOPSIS
        Proxy: forwards to the real Get-SPOBrowserIdleSignOut running in the
        SharePoint child process.
    #>
    [CmdletBinding()]
    param()
    $r = Invoke-BaselineSpoChildCommand -Command 'GetIdleSignOut'
    return [pscustomobject]@{
        Enabled     = $r.Enabled
        WarnAfter   = [timespan]::FromMinutes($r.WarnAfterMinutes)
        SignOutAfter = [timespan]::FromMinutes($r.SignOutAfterMinutes)
    }
}

function Set-SPOBrowserIdleSignOut {
    <#
    .SYNOPSIS
        Proxy: forwards to the real Set-SPOBrowserIdleSignOut running in the
        SharePoint child process.
    #>
    [CmdletBinding()]
    param(
        [Parameter()] [bool]$Enabled,
        [Parameter()] [timespan]$WarnAfter,
        [Parameter()] [timespan]$SignOutAfter
    )
    Invoke-BaselineSpoChildCommand -Command 'SetIdleSignOut' -Params @{
        Enabled             = $Enabled
        WarnAfterMinutes    = [int]$WarnAfter.TotalMinutes
        SignOutAfterMinutes = [int]$SignOutAfter.TotalMinutes
    } | Out-Null
}

function Disconnect-SPOService {
    <#
    .SYNOPSIS
        Proxy: gracefully shuts down the SharePoint child process. Called by
        BaselineCore.psm1's (unmodified) Disconnect-BaselineWorkload, which
        the app-only entry point reuses as-is.
    #>
    [CmdletBinding()]
    param()
    Stop-BaselineSpoChildProcess
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
        connects only to the services listed in -Services.

        Graph, ExchangeOnline, and Teams all receive that same resolved
        certificate object directly. SharePointOnline is the one exception:
        it's connected via a separate native Windows PowerShell 5.1 child
        process instead (see this file's header comment for why) - the
        resolved certificate is exported to a short-lived temp .pfx to hand
        to that process, never passed as a live object.

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
                    Connect-BaselineSpoServiceViaChildProcess -Url $SpoAdminUrl -AppId $AppId -TenantId $TenantId -Certificate $cert
                }
            }
        }
        catch {
            throw "Failed to connect to $service (app-only, certificate input '$($PSCmdlet.ParameterSetName)'): $($_.Exception.Message). This is one of three distinct problems, each with a different fix - see README.md's 'App-only (certificate) authentication' section, 'Troubleshooting a connection failure' subsection: (1) the certificate has expired, was revoked, or doesn't match what's uploaded to the app registration; (2) the required API permission for $service was not granted or not admin-consented; (3) $service additionally requires a directory role assignment on the app's service principal (Exchange Administrator for ExchangeOnline, Teams Administrator for Teams, SharePoint Administrator for SharePointOnline) that has not been made, separate from API permissions."
        }
        Set-BaselineWorkloadConnectedState -Connection $service -Connected $true
    }
}

Export-ModuleMember -Function @(
    'Connect-M365BaselineServicesAppOnly',
    'Get-SPOTenant',
    'Set-SPOTenant',
    'Get-SPOBrowserIdleSignOut',
    'Set-SPOBrowserIdleSignOut',
    'Disconnect-SPOService'
)
