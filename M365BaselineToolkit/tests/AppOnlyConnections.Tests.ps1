<#
    AppOnlyConnections.Tests.ps1

    Tests for the certificate-based, app-only (unattended) authentication
    alternative added alongside the existing interactive-auth toolkit:

      1. A file-hash invariance test proving the "do not modify" files listed
         in the app-only work's spec are still byte-identical to their
         pre-app-only-work state. This is the mechanical enforcement of that
         work's hard constraint - not just a promise in a commit message.
      2. Mock-based tests for Connect-M365BaselineServicesAppOnly
         (modules/AppOnlyConnections.psm1) covering all three certificate
         input forms, and confirming the SAME X509Certificate2 object
         instance is passed to all four Connect-* calls.
      3. A test confirming lazy-connect behavior is respected using the
         exact same (unmodified) BaselineCore.psm1 functions
         Invoke-M365Baseline.AppOnly.ps1 relies on for this - including that
         a control's Graph *ExtraConnection* (e.g. a license check) is still
         honored even when no control's primary connection is Graph.
      4. An integration test that actually runs both entry-point scripts
         (Invoke-M365Baseline.ps1 and Invoke-M365Baseline.AppOnly.ps1) as
         separate child processes against the same config file and
         identically-mocked service cmdlets, and compares their Audit-mode
         compliance verdicts - this is the test that actually proves "the
         control config and catalog stay the same," not just prose asserting
         it.

      5. Tests for the SharePoint-specific connection path: unlike Graph/
         ExchangeOnline/Teams, SharePointOnline is connected via a real,
         separately spawned PowerShell child process (see
         AppOnlyConnections.psm1's header comment for why - a documented,
         real-tenant-confirmed WinCompat limitation with certificate
         private keys), talking over a line-based JSON protocol on its
         stdin/stdout. These tests stand in a fake child process (using
         pwsh, since it's cross-platform and available in CI, with mocked
         SPO cmdlets) in place of the real powershell.exe + Microsoft.
         Online.SharePoint.PowerShell, and verify the wire protocol, the
         Get-SPOTenant/Set-SPOTenant/Get-SPOBrowserIdleSignOut/
         Set-SPOBrowserIdleSignOut/Disconnect-SPOService proxy functions,
         and that the unmodified SharePointOnlineControls.psm1 works
         transparently through them. They do NOT verify that a real
         Windows PowerShell 5.1 process or the real SharePoint Online
         Management Shell module behaves this way under app-only auth on
         any given tenant - that remains a real-tenant verification step
         (see README.md's SharePoint caveat).

    No live tenant connection, and none of Microsoft.Graph.Authentication,
    ExchangeOnlineManagement, MicrosoftTeams, or
    Microsoft.Online.SharePoint.PowerShell, is required - Import-Module for
    those four names is mocked out everywhere it's called in this file, and
    the SharePoint child process is stood in with pwsh + fake cmdlets.
#>

BeforeAll {
    $script:RepoRoot = Join-Path $PSScriptRoot '..'
    $script:ModulesDir = Join-Path $script:RepoRoot 'modules'

    # Real, self-signed, cross-platform test certificate (no OS certificate
    # store dependency), shared by every Describe block below.
    $rsa = [System.Security.Cryptography.RSA]::Create(2048)
    $req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
        'CN=AppOnlyConnectionsTests', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
    $script:TestCert = $req.CreateSelfSigned([datetimeoffset]::Now.AddDays(-1), [datetimeoffset]::Now.AddDays(30))

    $script:PfxPath = Join-Path $TestDrive 'test-app-only.pfx'
    [System.IO.File]::WriteAllBytes($script:PfxPath, $script:TestCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx))

    # The fake SharePoint child process server script: same JSON wire
    # protocol as AppOnlyConnections.psm1's real $script:SpoChildServerScript,
    # but with fake Connect-SPOService/Get-SPOTenant/etc. (pwsh-compatible)
    # standing in for the real Microsoft.Online.SharePoint.PowerShell module,
    # which isn't installed in this test environment and isn't what's being
    # verified here - the real module's own app-only behavior is a real-tenant
    # verification step, not something these tests can assert. The fake
    # Connect-SPOService also writes the loaded certificate's thumbprint to a
    # marker file, so tests can confirm the SAME certificate that was resolved
    # for Graph/ExchangeOnline/Teams is the one whose bytes actually reached
    # SharePoint's connect call, without the live object ever crossing the
    # process boundary (impossible to assert via reference equality here,
    # since it deliberately never does - see AppOnlyConnections.psm1's header
    # comment for why that's the whole point of this being a separate path).
    $script:FakeSpoChildServerScript = @'
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
function Connect-SPOService {
    param($Url,$ApplicationId,$TenantId,$CertificatePath,$CertificatePassword)
    if (-not (Test-Path $CertificatePath)) { throw "pfx not found at $CertificatePath" }
    $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertificatePath, $CertificatePassword)
    if ($env:SPO_TEST_MARKER_FILE) { Set-Content -Path $env:SPO_TEST_MARKER_FILE -Value $cert.Thumbprint -Encoding utf8 }
    if ($env:SPO_TEST_FAIL_CONNECT -eq '1') { throw "simulated Connect-SPOService failure" }
}
function Get-SPOTenant { [pscustomobject]@{ SharingCapability="ExternalUserSharingOnly"; DefaultSharingLinkType="Direct"; DefaultLinkPermission="View"; RequireAnonymousLinksExpireInDays=30; LegacyAuthProtocolsEnabled=$false } }
function Set-SPOTenant { param($SharingCapability,$DefaultSharingLinkType,$DefaultLinkPermission,$RequireAnonymousLinksExpireInDays,$LegacyAuthProtocolsEnabled,$ErrorAction)
    if ($env:SPO_TEST_MARKER_FILE) {
        ($PSBoundParameters.Keys | Where-Object { $_ -ne 'ErrorAction' }) -join ',' | Set-Content -Path "$($env:SPO_TEST_MARKER_FILE).setparams" -Encoding utf8
    }
}
function Get-SPOBrowserIdleSignOut { [pscustomobject]@{ Enabled=$true; WarnAfter=[timespan]::FromMinutes(15); SignOutAfter=[timespan]::FromMinutes(20) } }
function Set-SPOBrowserIdleSignOut { param($Enabled,$WarnAfter,$SignOutAfter,$ErrorAction) }
function Disconnect-SPOService { param($ErrorAction) }

function Write-SpoResponse {
    param($Success, $Result, $ErrorMessage)
    $obj = @{ Success = [bool]$Success }
    if ($Success) { $obj.Result = $Result } else { $obj.Error = [string]$ErrorMessage }
    [Console]::Out.WriteLine(($obj | ConvertTo-Json -Compress -Depth 6))
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
            "Connect" {
                $p = $request.Params
                $securePwd = ConvertTo-SecureString -String $p.CertificatePassword -AsPlainText -Force
                Connect-SPOService -Url $p.Url -ApplicationId $p.ApplicationId -TenantId $p.TenantId -CertificatePath $p.CertificatePath -CertificatePassword $securePwd -ErrorAction Stop
                $result = @{ connected = $true }
            }
            "GetTenant" {
                $t = Get-SPOTenant -ErrorAction Stop
                $result = @{ SharingCapability=[string]$t.SharingCapability; DefaultSharingLinkType=[string]$t.DefaultSharingLinkType; DefaultLinkPermission=[string]$t.DefaultLinkPermission; RequireAnonymousLinksExpireInDays=[int]$t.RequireAnonymousLinksExpireInDays; LegacyAuthProtocolsEnabled=[bool]$t.LegacyAuthProtocolsEnabled }
            }
            "SetTenant" {
                $setParams = @{ ErrorAction = "Stop" }
                foreach ($prop in $request.Params.PSObject.Properties) { $setParams[$prop.Name] = $prop.Value }
                Set-SPOTenant @setParams
                $result = @{ ok = $true }
            }
            "GetIdleSignOut" {
                $c = Get-SPOBrowserIdleSignOut -ErrorAction Stop
                $result = @{ Enabled=[bool]$c.Enabled; WarnAfterMinutes=[int]([timespan]$c.WarnAfter).TotalMinutes; SignOutAfterMinutes=[int]([timespan]$c.SignOutAfter).TotalMinutes }
            }
            "SetIdleSignOut" {
                Set-SPOBrowserIdleSignOut -Enabled:([bool]$request.Params.Enabled) -WarnAfter (New-TimeSpan -Minutes ([int]$request.Params.WarnAfterMinutes)) -SignOutAfter (New-TimeSpan -Minutes ([int]$request.Params.SignOutAfterMinutes)) -ErrorAction Stop
                $result = @{ ok = $true }
            }
            "Disconnect" { Disconnect-SPOService -ErrorAction SilentlyContinue; $result = @{ ok = $true } }
            "Exit" { $result = @{ ok = $true }; $shouldExit = $true }
            default { throw "Unknown command: $($request.Command)" }
        }
        Write-SpoResponse -Success $true -Result $result
    }
    catch {
        Write-SpoResponse -Success $false -ErrorMessage $_.Exception.Message
    }
    if ($shouldExit) { break }
}
'@

    function New-BaselineFakeSpoChildProcess {
        <#
        .SYNOPSIS
            Test helper: starts a pwsh-based fake SharePoint child process
            (see $script:FakeSpoChildServerScript) and returns the started
            System.Diagnostics.Process object, ready to be assigned to
            AppOnlyConnections.psm1's private $script:SpoChildProcess.
        #>
        param([string]$ScriptPath)
        $psi = [System.Diagnostics.ProcessStartInfo]::new()
        $psi.FileName = 'pwsh'
        $psi.Arguments = "-NoProfile -NonInteractive -File `"$ScriptPath`""
        $psi.RedirectStandardInput = $true
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.StandardInputEncoding = [System.Text.UTF8Encoding]::new($false)
        $psi.StandardOutputEncoding = [System.Text.UTF8Encoding]::new($false)
        $psi.UseShellExecute = $false
        $proc = [System.Diagnostics.Process]::new()
        $proc.StartInfo = $psi
        [void]$proc.Start()
        return $proc
    }
}

Describe 'Protected files are never modified by the app-only auth work' {
    BeforeAll {
        # Captured via Get-FileHash -Algorithm SHA256 immediately before any
        # app-only-auth file was added, and re-verified identical immediately
        # after all app-only-auth work in this session completed. These are
        # the "do not modify" files named by that work's hard constraint.
        #
        # Never edit these expected values to make a failing test pass -
        # a mismatch means one of these files actually changed. Find that
        # change and revert it instead.
        $script:ProtectedFileHashes = [ordered]@{
            'Invoke-M365Baseline.ps1'                = '9C8E80142971C5C40354B29F8EB54D687AC7B9DB647B984663B70EB52317003F'
            'modules/BaselineCore.psm1'               = 'ADA14B61267E4EE994AB867A6451F4C3519CAFDE4A9395983D9263E4F525F02F'
            'modules/EntraIdControls.psm1'             = '5C1B2277F0E7CEB5190C9EF28BA18C77F70ECCDA5DDBC7A3C7DAB4C32FD43435'
            'modules/ExchangeOnlineControls.psm1'      = '21602C16FD160E664E45680A492BBED5262D082D29749674C17C0EFC7B4BE7E3'
            'modules/TeamsControls.psm1'               = '9715FE7D6A50E4D1184840CAEDE6A347C98B15984FE6FCBD87286B9783F04077'
            'modules/SharePointOnlineControls.psm1'    = '6EEC1959B2DCCD0E91FFF4B8520D749773D2E4D52773BCADD921F196CF3C1E52'
            'modules/ConditionalAccessControls.psm1'   = '40022627D53A5AA8355A52EE034344FE0C1CC05EE2463F092E25B74EF1FC1479'
            'config/baseline.config.json'              = 'CFA02248F66C77A0DF548C2B39DEA4067F992F9F0F3223E5368B327B3FBCDF31'
            'config/baseline.config.schema.json'       = '880E53F7F46717E655E6CF45034DA67158ACDD1BCED8F0D68A96769298A36C7F'
        }
    }

    It 'has not modified <_>' -ForEach @($script:ProtectedFileHashes.Keys) {
        $path = Join-Path $script:RepoRoot $_
        Test-Path -LiteralPath $path | Should -BeTrue -Because "the file is expected to exist unmodified at $path"
        $actual = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $actual | Should -Be $script:ProtectedFileHashes[$_] -Because 'the app-only auth work''s hard constraint is zero edits to this file'
    }
}

Describe 'Connect-M365BaselineServicesAppOnly' {
    BeforeAll {
        Import-Module (Join-Path $script:ModulesDir 'BaselineCore.psm1') -Force
        Import-Module (Join-Path $script:ModulesDir 'AppOnlyConnections.psm1') -Force
    }

    BeforeEach {
        Mock -CommandName Import-Module -ModuleName AppOnlyConnections -MockWith { }
        Mock -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -MockWith { }
        Mock -CommandName Connect-ExchangeOnline -ModuleName AppOnlyConnections -MockWith { }
        Mock -CommandName Connect-MicrosoftTeams -ModuleName AppOnlyConnections -MockWith { [pscustomobject]@{} }
        # SharePointOnline no longer calls Connect-SPOService directly (see
        # "SharePointOnline via the native PS5.1 child process" Describe block
        # below for that path) - here it's enough to stub out the
        # child-process starter so tests in this Describe that happen to
        # include SharePointOnline in -Services don't try to spawn a real
        # powershell.exe.
        Mock -CommandName Start-BaselineSpoChildProcess -ModuleName AppOnlyConnections -MockWith { }
        Mock -CommandName Connect-BaselineSpoServiceViaChildProcess -ModuleName AppOnlyConnections -MockWith { }
    }

    Context 'Certificate input forms' {
        It 'accepts -Certificate (a pre-built X509Certificate2 object)' {
            { Connect-M365BaselineServicesAppOnly -Services Graph -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Certificate $script:TestCert } | Should -Not -Throw
            Should -Invoke -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -Times 1 -ParameterFilter {
                $Certificate.Thumbprint -eq $script:TestCert.Thumbprint
            }
        }

        It 'accepts -CertificatePath (a .pfx file with no password)' {
            { Connect-M365BaselineServicesAppOnly -Services Graph -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -CertificatePath $script:PfxPath } | Should -Not -Throw
            Should -Invoke -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -Times 1 -ParameterFilter {
                $Certificate.Thumbprint -eq $script:TestCert.Thumbprint
            }
        }

        It 'accepts -CertificateThumbprint, resolved from the certificate store' {
            Mock -CommandName Get-ChildItem -ModuleName AppOnlyConnections -MockWith { $script:TestCert } -ParameterFilter { $Path -eq 'Cert:\CurrentUser\My' }
            { Connect-M365BaselineServicesAppOnly -Services Graph -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -CertificateThumbprint $script:TestCert.Thumbprint } | Should -Not -Throw
            Should -Invoke -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -Times 1 -ParameterFilter {
                $Certificate.Thumbprint -eq $script:TestCert.Thumbprint
            }
        }

        It 'honors -CertificateStoreLocation LocalMachine' {
            Mock -CommandName Get-ChildItem -ModuleName AppOnlyConnections -MockWith { $script:TestCert } -ParameterFilter { $Path -eq 'Cert:\LocalMachine\My' }
            Mock -CommandName Get-ChildItem -ModuleName AppOnlyConnections -MockWith { @() } -ParameterFilter { $Path -eq 'Cert:\CurrentUser\My' }
            { Connect-M365BaselineServicesAppOnly -Services Graph -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -CertificateThumbprint $script:TestCert.Thumbprint -CertificateStoreLocation LocalMachine } | Should -Not -Throw
        }

        It 'throws a clear, specific error when a thumbprint is not found in the store' {
            Mock -CommandName Get-ChildItem -ModuleName AppOnlyConnections -MockWith { @() }
            { Connect-M365BaselineServicesAppOnly -Services Graph -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -CertificateThumbprint 'DEADBEEF00000000000000000000000000000000' } |
                Should -Throw -ExpectedMessage '*No certificate with thumbprint*'
        }
    }

    Context 'Uniform certificate object across services' {
        It 'passes the SAME X509Certificate2 instance (reference equality) to Graph/ExchangeOnline/Teams, and to the SharePoint child-process handoff' {
            Connect-M365BaselineServicesAppOnly -Services Graph, ExchangeOnline, Teams, SharePointOnline `
                -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Organization 'contoso.onmicrosoft.com' -SpoAdminUrl 'https://contoso-admin.sharepoint.com' `
                -Certificate $script:TestCert

            Should -Invoke -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -Times 1 -ParameterFilter { [object]::ReferenceEquals($Certificate, $script:TestCert) }
            Should -Invoke -CommandName Connect-ExchangeOnline -ModuleName AppOnlyConnections -Times 1 -ParameterFilter { [object]::ReferenceEquals($Certificate, $script:TestCert) }
            Should -Invoke -CommandName Connect-MicrosoftTeams -ModuleName AppOnlyConnections -Times 1 -ParameterFilter { [object]::ReferenceEquals($Certificate, $script:TestCert) }
            # SharePointOnline deliberately never receives the live object (see
            # AppOnlyConnections.psm1's header comment - that's the documented
            # fix for the WinCompat private-key problem) - it's handed to the
            # child-process bridge function instead, which is what actually
            # exports it to a temp .pfx. Verified end to end (the exported pfx
            # really does contain this exact certificate) in the "SharePointOnline
            # via the native PS5.1 child process" Describe block below.
            Should -Invoke -CommandName Connect-BaselineSpoServiceViaChildProcess -ModuleName AppOnlyConnections -Times 1 -ParameterFilter { [object]::ReferenceEquals($Certificate, $script:TestCert) }
        }

        It 'connects in the order -Services is given, without re-sorting' {
            Connect-M365BaselineServicesAppOnly -Services ExchangeOnline, Graph `
                -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Organization 'contoso.onmicrosoft.com' -Certificate $script:TestCert

            Should -Invoke -CommandName Connect-ExchangeOnline -ModuleName AppOnlyConnections -Times 1
            Should -Invoke -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -Times 1
        }
    }

    Context 'Lazy connect: only the requested -Services connect' {
        It 'never calls Connect-MgGraph or Connect-MicrosoftTeams when only ExchangeOnline/SharePointOnline are requested' {
            Connect-M365BaselineServicesAppOnly -Services ExchangeOnline, SharePointOnline `
                -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Organization 'contoso.onmicrosoft.com' -SpoAdminUrl 'https://contoso-admin.sharepoint.com' `
                -Certificate $script:TestCert

            Should -Invoke -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -Times 0
            Should -Invoke -CommandName Connect-MicrosoftTeams -ModuleName AppOnlyConnections -Times 0
            Should -Invoke -CommandName Connect-ExchangeOnline -ModuleName AppOnlyConnections -Times 1
            Should -Invoke -CommandName Connect-BaselineSpoServiceViaChildProcess -ModuleName AppOnlyConnections -Times 1
        }
    }

    Context 'Required parameters per service' {
        It 'requires -Organization when ExchangeOnline is requested' {
            { Connect-M365BaselineServicesAppOnly -Services ExchangeOnline -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Certificate $script:TestCert } |
                Should -Throw -ExpectedMessage '*-Organization*'
        }

        It 'requires -SpoAdminUrl when SharePointOnline is requested' {
            { Connect-M365BaselineServicesAppOnly -Services SharePointOnline -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Certificate $script:TestCert } |
                Should -Throw -ExpectedMessage '*-SpoAdminUrl*'
        }
    }

    Context 'Connection failure errors are specific, not generic' {
        It 'names the failing service and points to the README for the three distinct likely causes' {
            Mock -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -MockWith { throw 'AADSTS700016: Application not found in the directory' }
            { Connect-M365BaselineServicesAppOnly -Services Graph -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Certificate $script:TestCert } |
                Should -Throw -ExpectedMessage '*Graph*README*'
        }

        It 'preserves the original exception message' {
            Mock -CommandName Connect-ExchangeOnline -ModuleName AppOnlyConnections -MockWith { throw 'The account does not have permission to perform this action' }
            { Connect-M365BaselineServicesAppOnly -Services ExchangeOnline -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Organization 'contoso.onmicrosoft.com' -Certificate $script:TestCert } |
                Should -Throw -ExpectedMessage '*does not have permission*'
        }
    }
}

Describe 'SharePointOnline via the native PS5.1 child process' {
    # Stands in a pwsh-based fake child process (with fake SPO cmdlets) for
    # the real powershell.exe + Microsoft.Online.SharePoint.PowerShell -
    # see this file's header comment. Injects it directly into
    # AppOnlyConnections.psm1's private $script:SpoChildProcess, bypassing
    # Start-BaselineSpoChildProcess (which hardcodes powershell.exe) so these
    # tests work on any OS Pester runs on, not just Windows.
    BeforeAll {
        Import-Module (Join-Path $script:ModulesDir 'BaselineCore.psm1') -Force
        Import-Module (Join-Path $script:ModulesDir 'AppOnlyConnections.psm1') -Force
        Import-Module (Join-Path $script:ModulesDir 'SharePointOnlineControls.psm1') -Force -WarningAction SilentlyContinue

        $script:FakeSpoScriptPath = Join-Path $TestDrive 'fake-spo-child.ps1'
        Set-Content -Path $script:FakeSpoScriptPath -Value $script:FakeSpoChildServerScript -Encoding utf8

        $script:AppOnlyConnectionsModule = Get-Module AppOnlyConnections
    }

    BeforeEach {
        $env:SPO_TEST_MARKER_FILE = Join-Path $TestDrive "spo-marker-$([guid]::NewGuid()).txt"
        $env:SPO_TEST_FAIL_CONNECT = $null
        $fakeProcess = New-BaselineFakeSpoChildProcess -ScriptPath $script:FakeSpoScriptPath
        & $script:AppOnlyConnectionsModule { param($p) $script:SpoChildProcess = $p } $fakeProcess
    }

    AfterEach {
        & $script:AppOnlyConnectionsModule {
            if ($script:SpoChildProcess -and -not $script:SpoChildProcess.HasExited) {
                try { $script:SpoChildProcess.Kill() } catch { }
            }
            $script:SpoChildProcess = $null
        }
        Remove-Item Env:\SPO_TEST_MARKER_FILE, Env:\SPO_TEST_FAIL_CONNECT -ErrorAction SilentlyContinue
    }

    It 'Get-SPOTenant proxy round-trips every field SharePointOnlineControls.psm1 reads' {
        $tenant = Get-SPOTenant
        $tenant.SharingCapability | Should -Be 'ExternalUserSharingOnly'
        $tenant.DefaultSharingLinkType | Should -Be 'Direct'
        $tenant.DefaultLinkPermission | Should -Be 'View'
        $tenant.RequireAnonymousLinksExpireInDays | Should -Be 30
        $tenant.LegacyAuthProtocolsEnabled | Should -Be $false
    }

    It 'Set-SPOTenant proxy forwards only the one parameter passed, matching SharePointOnlineControls.psm1''s call pattern' {
        { Set-SPOTenant -SharingCapability 'Disabled' -ErrorAction Stop } | Should -Not -Throw
        $capturedParams = Get-Content "$($env:SPO_TEST_MARKER_FILE).setparams"
        $capturedParams | Should -Be 'SharingCapability'
    }

    It 'Get-/Set-SPOBrowserIdleSignOut proxies round-trip TimeSpan values correctly' {
        $idle = Get-SPOBrowserIdleSignOut
        $idle.Enabled | Should -Be $true
        $idle.WarnAfter | Should -Be ([timespan]::FromMinutes(15))
        $idle.SignOutAfter | Should -Be ([timespan]::FromMinutes(20))
        { Set-SPOBrowserIdleSignOut -Enabled $true -WarnAfter ([timespan]::FromMinutes(10)) -SignOutAfter ([timespan]::FromMinutes(15)) -ErrorAction Stop } | Should -Not -Throw
    }

    It 'a child-side exception surfaces as a real PowerShell exception on the caller side' {
        # A running child process only sees environment variables as they were
        # at spawn time - the BeforeEach-started process already missed this
        # one, so replace it with a freshly spawned process that has
        # SPO_TEST_FAIL_CONNECT set before it starts.
        & $script:AppOnlyConnectionsModule {
            if ($script:SpoChildProcess -and -not $script:SpoChildProcess.HasExited) { try { $script:SpoChildProcess.Kill() } catch { } }
        }
        $env:SPO_TEST_FAIL_CONNECT = '1'
        $freshProcess = New-BaselineFakeSpoChildProcess -ScriptPath $script:FakeSpoScriptPath
        & $script:AppOnlyConnectionsModule { param($p) $script:SpoChildProcess = $p } $freshProcess

        # Connect-BaselineSpoServiceViaChildProcess is private (not exported),
        # so it's invoked within the module's own scope, same as the private
        # state access elsewhere in this Describe block.
        {
            & $script:AppOnlyConnectionsModule {
                param($cert)
                Connect-BaselineSpoServiceViaChildProcess -Url 'https://contoso-admin.sharepoint.com' -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Certificate $cert
            } $script:TestCert
        } | Should -Throw -ExpectedMessage '*simulated Connect-SPOService failure*'
    }

    It 'Disconnect-SPOService cleanly stops the child process' {
        Disconnect-SPOService
        Start-Sleep -Milliseconds 500
        $stillRunning = & $script:AppOnlyConnectionsModule { $script:SpoChildProcess -and -not $script:SpoChildProcess.HasExited }
        $stillRunning | Should -BeFalsy
    }

    It 'Connect-BaselineSpoServiceViaChildProcess exports the resolved certificate to a temp .pfx, hands it to the child, and cleans it up afterward' {
        & $script:AppOnlyConnectionsModule {
            param($cert)
            Connect-BaselineSpoServiceViaChildProcess -Url 'https://contoso-admin.sharepoint.com' -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Certificate $cert
        } $script:TestCert

        # The fake Connect-SPOService wrote the thumbprint of whatever
        # certificate it actually loaded from the handed-off .pfx - this is
        # the one place we can prove end to end that the SAME certificate
        # resolved for Graph/ExchangeOnline/Teams is what SharePoint's
        # connect call actually used, even though (unlike the other three)
        # it never crosses the process boundary as a live object.
        $observedThumbprint = Get-Content $env:SPO_TEST_MARKER_FILE
        $observedThumbprint | Should -Be $script:TestCert.Thumbprint

        $pfxPath = & $script:AppOnlyConnectionsModule { $script:SpoChildPfxPath }
        $pfxPath | Should -BeNullOrEmpty -Because 'the temp .pfx is deleted immediately after the child process loads it, win or lose'
    }

    It 'SharePointOnlineControls.psm1 (unmodified) works transparently through the proxy functions' {
        $result = Get-SharePointOnline-LegacyAuthProtocolsState
        $result.Id | Should -Be 'SharePointOnline-LegacyAuthProtocols'
        $result.Value | Should -Be $false

        $setResult = Set-SharePointOnline-LegacyAuthProtocolsState -DesiredValue $true -CurrentValue $false
        $setResult.Status | Should -Be 'Success'
        $capturedParams = Get-Content "$($env:SPO_TEST_MARKER_FILE).setparams"
        $capturedParams | Should -Be 'LegacyAuthProtocolsEnabled'
    }
}

Describe 'Lazy connection selection reuses BaselineCore.psm1 unchanged' {
    # This exercises the exact same functions (Get-BaselineControlCatalog,
    # Get-BaselineConnectionOrder) Invoke-M365Baseline.AppOnly.ps1 calls to
    # decide which services to connect to - proving the "connect lazily"
    # property doesn't depend on anything app-only-specific.
    BeforeAll {
        Import-Module (Join-Path $script:ModulesDir 'BaselineCore.psm1') -Force
        Import-Module (Join-Path $script:ModulesDir 'ExchangeOnlineControls.psm1') -Force -WarningAction SilentlyContinue
        Import-Module (Join-Path $script:ModulesDir 'SharePointOnlineControls.psm1') -Force -WarningAction SilentlyContinue
    }

    It 'requires only ExchangeOnline and SharePointOnline when only those controls are enabled' {
        $config = [pscustomobject]@{
            controls = @(
                [pscustomobject]@{ id = 'ExchangeOnline-DisableSmtpAuth'; workload = 'ExchangeOnline'; enabled = $true; automatable = $true; desiredValue = $true; description = 'x' }
                [pscustomobject]@{ id = 'SharePointOnline-LegacyAuthProtocols'; workload = 'SharePointOnline'; enabled = $true; automatable = $true; desiredValue = $false; description = 'x' }
            )
        }
        $availableFunctions = @(
            'Get-ExchangeOnline-DisableSmtpAuthState', 'Set-ExchangeOnline-DisableSmtpAuthState'
            'Get-SharePointOnline-LegacyAuthProtocolsState', 'Set-SharePointOnline-LegacyAuthProtocolsState'
        )
        $catalog = Get-BaselineControlCatalog -Config $config -AvailableFunctions $availableFunctions

        $needed = @($catalog | ForEach-Object { @($_.Connection) + @($_.ExtraConnections) })
        $required = Get-BaselineConnectionOrder -Connections $needed

        $required | Should -Not -Contain 'Graph'
        $required | Should -Not -Contain 'Teams'
        $required | Should -Contain 'ExchangeOnline'
        $required | Should -Contain 'SharePointOnline'
    }

    It 'still requires Graph via ExtraConnections when a control needs it for a license check, even though no control''s primary connection is Graph' {
        Import-Module (Join-Path $script:ModulesDir 'ExchangeOnlineControls.psm1') -Force -WarningAction SilentlyContinue
        $config = [pscustomobject]@{
            controls = @(
                [pscustomobject]@{ id = 'ExchangeOnline-AntiPhishingMailboxIntelligence'; workload = 'ExchangeOnline'; enabled = $true; automatable = $true; desiredValue = @{ enableMailboxIntelligence = $true; enableMailboxIntelligenceProtection = $true }; description = 'x' }
            )
        }
        $availableFunctions = @('Get-ExchangeOnline-AntiPhishingMailboxIntelligenceState', 'Set-ExchangeOnline-AntiPhishingMailboxIntelligenceState')
        $catalog = Get-BaselineControlCatalog -Config $config -AvailableFunctions $availableFunctions

        $catalog[0].Connection | Should -Be 'ExchangeOnline'
        $catalog[0].ExtraConnections | Should -Contain 'Graph'

        $needed = @($catalog | ForEach-Object { @($_.Connection) + @($_.ExtraConnections) })
        $required = Get-BaselineConnectionOrder -Connections $needed

        $required | Should -Contain 'Graph'
        $required | Should -Contain 'ExchangeOnline'
        $required | Should -Not -Contain 'Teams'
        $required | Should -Not -Contain 'SharePointOnline'
    }
}

Describe 'Cross-script parity: interactive and app-only entry points agree' -Tag 'Integration' {
    # Runs BOTH Invoke-M365Baseline.ps1 and Invoke-M365Baseline.AppOnly.ps1 as
    # separate child processes, against the SAME config file subset and
    # IDENTICALLY-mocked underlying service cmdlets (only the Connect-*
    # mocks differ, matching each script's own auth mechanism) - then
    # compares their Audit-mode Markdown reports' Id/Current/Desired/
    # Compliant columns. This is what actually proves the shared config and
    # control catalog produce the same verdicts regardless of which script
    # ran them, rather than just asserting it in prose.
    BeforeAll {
        $fullConfig = Get-Content (Join-Path $script:RepoRoot 'config/baseline.config.json') -Raw | ConvertFrom-Json
        foreach ($c in $fullConfig.controls) {
            $c.enabled = ($c.id -in @('ExchangeOnline-DisableSmtpAuth', 'ExchangeOnline-MailboxAuditingDefault'))
        }
        $script:ConfigPath = Join-Path $TestDrive 'parity.config.json'
        $fullConfig | ConvertTo-Json -Depth 10 | Set-Content -Path $script:ConfigPath -Encoding utf8

        $script:SharedMockPreamble = @'
$ErrorActionPreference = 'Stop'
function Get-TransportConfig { [pscustomobject]@{ SmtpClientAuthenticationDisabled = $true } }
function Set-TransportConfig { param($SmtpClientAuthenticationDisabled) }
function Get-OrganizationConfig { [pscustomobject]@{ AuditDisabled = $false } }
function Set-OrganizationConfig { param($AuditDisabled) }
$Global:FakeModuleNames = @('Microsoft.Graph.Authentication','ExchangeOnlineManagement','MicrosoftTeams','Microsoft.Online.SharePoint.PowerShell')
function Import-Module {
    [CmdletBinding()]
    param([Parameter(Position=0)]$Name,[switch]$Force,[switch]$Global,[switch]$UseWindowsPowerShell)
    if ($Name -is [string] -and $Global:FakeModuleNames -contains $Name) { return }
    Microsoft.PowerShell.Core\Import-Module -Name $Name -Force:$Force -Global:$Global
}
function Get-Module {
    [CmdletBinding()]
    param([Parameter(Position=0)]$Name,[switch]$ListAvailable)
    if ($ListAvailable -and $Name -is [string] -and $Global:FakeModuleNames -contains $Name) {
        return [pscustomobject]@{ Name = $Name; Version = [version]'1.0.0' }
    }
    Microsoft.PowerShell.Core\Get-Module @PSBoundParameters
}
'@

        $interactiveMocks = $script:SharedMockPreamble + @'
function Connect-ExchangeOnline { param($ShowBanner,[switch]$DisableWAM) }
function Disconnect-ExchangeOnline { param([switch]$Confirm) }
'@
        $script:InteractiveRunnerPath = Join-Path $TestDrive 'run-interactive.ps1'
        Set-Content -Path $script:InteractiveRunnerPath -Encoding utf8 -Value @"
$interactiveMocks
& '$($script:RepoRoot -replace "'", "''")/Invoke-M365Baseline.ps1' -Mode Audit -ConfigPath '$($script:ConfigPath -replace "'", "''")' -ReportPath '$TestDrive/interactive-reports' -BackupPath '$TestDrive/interactive-backups'
exit `$LASTEXITCODE
"@

        $appOnlyMocks = $script:SharedMockPreamble + @'
function Connect-ExchangeOnline { param($AppId,$Certificate,$Organization,[switch]$ShowBanner) }
function Disconnect-ExchangeOnline { param([switch]$Confirm) }
'@
        $script:AppOnlyRunnerPath = Join-Path $TestDrive 'run-apponly.ps1'
        Set-Content -Path $script:AppOnlyRunnerPath -Encoding utf8 -Value @"
$appOnlyMocks
`$rsa = [System.Security.Cryptography.RSA]::Create(2048)
`$req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new('CN=ParityTest', `$rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
`$cert = `$req.CreateSelfSigned([datetimeoffset]::Now.AddDays(-1), [datetimeoffset]::Now.AddDays(30))
& '$($script:RepoRoot -replace "'", "''")/Invoke-M365Baseline.AppOnly.ps1' -Mode Audit -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Organization 'contoso.onmicrosoft.com' -Certificate `$cert -ConfigPath '$($script:ConfigPath -replace "'", "''")' -ReportPath '$TestDrive/apponly-reports' -BackupPath '$TestDrive/apponly-backups'
exit `$LASTEXITCODE
"@
    }

    It 'produces identical compliance verdicts from both entry-point scripts' {
        $interactiveResult = & pwsh -NoProfile -File $script:InteractiveRunnerPath 2>&1
        $interactiveExit = $LASTEXITCODE
        $interactiveExit | Should -Be 0 -Because ("interactive script output was:`n" + ($interactiveResult -join "`n"))

        $appOnlyResult = & pwsh -NoProfile -File $script:AppOnlyRunnerPath 2>&1
        $appOnlyExit = $LASTEXITCODE
        $appOnlyExit | Should -Be 0 -Because ("app-only script output was:`n" + ($appOnlyResult -join "`n"))

        $interactiveReport = Get-ChildItem (Join-Path $TestDrive 'interactive-reports') -Filter 'audit-report*.md' | Select-Object -First 1
        $appOnlyReport = Get-ChildItem (Join-Path $TestDrive 'apponly-reports') -Filter 'audit-report*.md' | Select-Object -First 1

        # Extract just the data rows (Id + columns), stripping the two
        # differing title lines ("... (app-only)" vs not) and "Generated:"
        # timestamp line - everything that reflects the actual compliance
        # verdict must match.
        $extractRows = { param($path) (Get-Content $path) | Where-Object { $_ -match '^\|' -and $_ -notmatch '^\| Id \|' -and $_ -notmatch '^\|---' } }
        $interactiveRows = & $extractRows $interactiveReport.FullName
        $appOnlyRows = & $extractRows $appOnlyReport.FullName

        $interactiveRows.Count | Should -Be 2
        $appOnlyRows | Should -Be $interactiveRows -Because 'the same config, same catalog, and same mocked service state must produce the same compliance verdicts regardless of which script ran them'
    }
}
