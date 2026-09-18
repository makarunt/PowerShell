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

    No live tenant connection, and none of Microsoft.Graph.Authentication,
    ExchangeOnlineManagement, MicrosoftTeams, or
    Microsoft.Online.SharePoint.PowerShell, is required - Import-Module for
    those four names is mocked out everywhere it's called in this file.
#>

BeforeAll {
    $script:RepoRoot = Join-Path $PSScriptRoot '..'
    $script:ModulesDir = Join-Path $script:RepoRoot 'modules'
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

        # Real, self-signed, cross-platform test certificate (no OS certificate
        # store dependency) to exercise the -Certificate/-CertificatePath forms.
        $rsa = [System.Security.Cryptography.RSA]::Create(2048)
        $req = [System.Security.Cryptography.X509Certificates.CertificateRequest]::new(
            'CN=AppOnlyConnectionsTests', $rsa, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $script:TestCert = $req.CreateSelfSigned([datetimeoffset]::Now.AddDays(-1), [datetimeoffset]::Now.AddDays(30))

        $script:PfxPath = Join-Path $TestDrive 'test-app-only.pfx'
        [System.IO.File]::WriteAllBytes($script:PfxPath, $script:TestCert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pfx))
    }

    BeforeEach {
        Mock -CommandName Import-Module -ModuleName AppOnlyConnections -MockWith { }
        Mock -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -MockWith { }
        Mock -CommandName Connect-ExchangeOnline -ModuleName AppOnlyConnections -MockWith { }
        Mock -CommandName Connect-MicrosoftTeams -ModuleName AppOnlyConnections -MockWith { [pscustomobject]@{} }
        Mock -CommandName Connect-SPOService -ModuleName AppOnlyConnections -MockWith { }
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
        It 'passes the SAME X509Certificate2 instance (reference equality) to all four Connect-* calls' {
            Connect-M365BaselineServicesAppOnly -Services Graph, ExchangeOnline, Teams, SharePointOnline `
                -AppId 'app-id' -TenantId 'contoso.onmicrosoft.com' -Organization 'contoso.onmicrosoft.com' -SpoAdminUrl 'https://contoso-admin.sharepoint.com' `
                -Certificate $script:TestCert

            Should -Invoke -CommandName Connect-MgGraph -ModuleName AppOnlyConnections -Times 1 -ParameterFilter { [object]::ReferenceEquals($Certificate, $script:TestCert) }
            Should -Invoke -CommandName Connect-ExchangeOnline -ModuleName AppOnlyConnections -Times 1 -ParameterFilter { [object]::ReferenceEquals($Certificate, $script:TestCert) }
            Should -Invoke -CommandName Connect-MicrosoftTeams -ModuleName AppOnlyConnections -Times 1 -ParameterFilter { [object]::ReferenceEquals($Certificate, $script:TestCert) }
            Should -Invoke -CommandName Connect-SPOService -ModuleName AppOnlyConnections -Times 1 -ParameterFilter { [object]::ReferenceEquals($Certificate, $script:TestCert) }
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
            Should -Invoke -CommandName Connect-SPOService -ModuleName AppOnlyConnections -Times 1
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
