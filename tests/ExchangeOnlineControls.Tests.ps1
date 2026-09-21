<#
    ExchangeOnlineControls.Tests.ps1

    Tests representative Exchange Online controls' Get-/Set- functions with
    mocked ExchangeOnlineManagement cmdlets. No live EXO connection or the
    ExchangeOnlineManagement module itself is required.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '../modules/ExchangeOnlineControls.psm1') -Force

    # ExchangeOnlineManagement isn't necessarily installed wherever this suite runs
    # (unlike Microsoft.Graph, which the other test files' mocks lean on existing for
    # real). Since Pester 5, Mock requires its target command to already be resolvable -
    # a real cmdlet, or an existing function - it can no longer "blindly" mock a name
    # that exists nowhere, the way Pester 4 did. Confirmed live: on a workstation with
    # Microsoft.Graph installed but not ExchangeOnlineManagement, every test in this
    # file failed with "CommandNotFoundException: Could not find Command
    # Get-OrganizationConfig" (and the same for every other EXO cmdlet below) before
    # Mock ever got a chance to run. These stubs exist purely so Mock has something to
    # find and replace - their own bodies are never reached once Mock -CommandName is
    # applied over them. Declared only when the real cmdlet isn't already present, so a
    # machine that DOES have ExchangeOnlineManagement installed keeps using the real
    # cmdlet's own (more accurate) parameter metadata for the mock instead.
    if (-not (Get-Command Get-OrganizationConfig -ErrorAction SilentlyContinue)) {
        function global:Get-OrganizationConfig { [CmdletBinding()] param() }
    }
    if (-not (Get-Command Set-OrganizationConfig -ErrorAction SilentlyContinue)) {
        function global:Set-OrganizationConfig { [CmdletBinding()] param([switch]$AuditDisabled) }
    }
    if (-not (Get-Command Get-AntiPhishPolicy -ErrorAction SilentlyContinue)) {
        function global:Get-AntiPhishPolicy { [CmdletBinding()] param([string]$Identity) }
    }
    if (-not (Get-Command Set-AntiPhishPolicy -ErrorAction SilentlyContinue)) {
        function global:Set-AntiPhishPolicy { [CmdletBinding()] param([string]$Identity, [switch]$EnableSpoofIntelligence, [switch]$EnableMailboxIntelligence, [switch]$EnableMailboxIntelligenceProtection) }
    }
    if (-not (Get-Command Get-TransportConfig -ErrorAction SilentlyContinue)) {
        function global:Get-TransportConfig { [CmdletBinding()] param() }
    }
    if (-not (Get-Command Set-TransportConfig -ErrorAction SilentlyContinue)) {
        function global:Set-TransportConfig { [CmdletBinding()] param([switch]$SmtpClientAuthenticationDisabled) }
    }
    if (-not (Get-Command Get-DkimSigningConfig -ErrorAction SilentlyContinue)) {
        function global:Get-DkimSigningConfig { [CmdletBinding()] param([string]$Identity) }
    }
    if (-not (Get-Command Set-DkimSigningConfig -ErrorAction SilentlyContinue)) {
        function global:Set-DkimSigningConfig { [CmdletBinding()] param([string]$Identity, [bool]$Enabled) }
    }
    if (-not (Get-Command New-DkimSigningConfig -ErrorAction SilentlyContinue)) {
        function global:New-DkimSigningConfig { [CmdletBinding()] param([string]$DomainName, [bool]$Enabled) }
    }
}

Describe 'ExchangeOnline-MailboxAuditingDefault' {

    Context 'Get-ExchangeOnline-MailboxAuditingDefaultState' {
        It 'inverts AuditDisabled to report the positive "auditing enabled" sense' {
            Mock -CommandName Get-OrganizationConfig -ModuleName ExchangeOnlineControls -MockWith {
                [pscustomobject]@{ AuditDisabled = $true }
            }
            (Get-ExchangeOnline-MailboxAuditingDefaultState).Value | Should -Be $false
        }
    }

    Context 'Set-ExchangeOnline-MailboxAuditingDefaultState' {
        It 'sets AuditDisabled to the inverse of the desired value when non-compliant' {
            Mock -CommandName Set-OrganizationConfig -ModuleName ExchangeOnlineControls -MockWith { }

            $result = Set-ExchangeOnline-MailboxAuditingDefaultState -DesiredValue $true -CurrentValue $false

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-OrganizationConfig -ModuleName ExchangeOnlineControls -Times 1 -ParameterFilter {
                $AuditDisabled -eq $false
            }
        }

        It 'is a no-op when already compliant' {
            Mock -CommandName Set-OrganizationConfig -ModuleName ExchangeOnlineControls -MockWith { }
            $result = Set-ExchangeOnline-MailboxAuditingDefaultState -DesiredValue $true -CurrentValue $true
            $result.Message | Should -Match 'Already compliant'
            Should -Invoke -CommandName Set-OrganizationConfig -ModuleName ExchangeOnlineControls -Times 0
        }
    }
}

Describe 'ExchangeOnline-DkimSigning' {

    Context 'Get-ExchangeOnline-DkimSigningState' {
        It 'lists every configured domain and its enabled state' {
            Mock -CommandName Get-DkimSigningConfig -ModuleName ExchangeOnlineControls -MockWith {
                param($Identity)
                if ($Identity) { return $null }
                @(
                    [pscustomobject]@{ Domain = 'contoso.com'; Enabled = $true }
                    [pscustomobject]@{ Domain = 'fabrikam.com'; Enabled = $false }
                )
            }
            $result = Get-ExchangeOnline-DkimSigningState
            $result.Value.domains.Count | Should -Be 2
            ($result.Value.domains | Where-Object domain -eq 'fabrikam.com').enabled | Should -Be $false
        }
    }

    Context 'Set-ExchangeOnline-DkimSigningState' {
        It 'throws an actionable error when desiredValue.domains is empty (no safe universal default)' {
            $desired = [pscustomobject]@{ domains = @() }
            { Set-ExchangeOnline-DkimSigningState -DesiredValue $desired -CurrentValue $null } | Should -Throw '*at least one domain*'
        }

        It 'creates a new signing config for a domain with none configured' {
            Mock -CommandName Get-DkimSigningConfig -ModuleName ExchangeOnlineControls -MockWith { param($Identity) $null }
            Mock -CommandName New-DkimSigningConfig -ModuleName ExchangeOnlineControls -MockWith { }
            Mock -CommandName Set-DkimSigningConfig -ModuleName ExchangeOnlineControls -MockWith { }

            $desired = [pscustomobject]@{ domains = @('contoso.com') }
            $result = Set-ExchangeOnline-DkimSigningState -DesiredValue $desired -CurrentValue ([pscustomobject]@{ domains = @() })

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName New-DkimSigningConfig -ModuleName ExchangeOnlineControls -Times 1 -ParameterFilter { $DomainName -eq 'contoso.com' }
            Should -Invoke -CommandName Set-DkimSigningConfig -ModuleName ExchangeOnlineControls -Times 0
        }

        It 'enables an existing but disabled signing config instead of creating a new one' {
            Mock -CommandName Get-DkimSigningConfig -ModuleName ExchangeOnlineControls -MockWith {
                param($Identity)
                [pscustomobject]@{ Domain = $Identity; Enabled = $false }
            }
            Mock -CommandName New-DkimSigningConfig -ModuleName ExchangeOnlineControls -MockWith { }
            Mock -CommandName Set-DkimSigningConfig -ModuleName ExchangeOnlineControls -MockWith { }

            $desired = [pscustomobject]@{ domains = @('contoso.com') }
            $result = Set-ExchangeOnline-DkimSigningState -DesiredValue $desired -CurrentValue ([pscustomobject]@{ domains = @() })

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-DkimSigningConfig -ModuleName ExchangeOnlineControls -Times 1 -ParameterFilter { $Identity -eq 'contoso.com' }
            Should -Invoke -CommandName New-DkimSigningConfig -ModuleName ExchangeOnlineControls -Times 0
        }
    }
}

Describe 'ExchangeOnline-DisableSmtpAuth' {

    Context 'Set-ExchangeOnline-DisableSmtpAuthState' {
        It 'disables SMTP AUTH when currently enabled' {
            Mock -CommandName Set-TransportConfig -ModuleName ExchangeOnlineControls -MockWith { }
            $result = Set-ExchangeOnline-DisableSmtpAuthState -DesiredValue $true -CurrentValue $false
            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-TransportConfig -ModuleName ExchangeOnlineControls -Times 1 -ParameterFilter { $SmtpClientAuthenticationDisabled -eq $true }
        }
    }
}

Describe 'ExchangeOnline-AntiPhishing (split into a base-EOP control and a Defender-for-O365-gated control)' {

    BeforeEach {
        # Re-imports with -Force on every test, exactly like Invoke-M365Baseline.ps1
        # does on every real run: BaselineCore's Get-BaselineSubscribedSkuCache is
        # module-scoped and would otherwise leak one test's mocked tenant license
        # into the next.
        Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
        Import-Module (Join-Path $PSScriptRoot '../modules/ExchangeOnlineControls.psm1') -Force

        Mock -CommandName Get-AntiPhishPolicy -ModuleName ExchangeOnlineControls -MockWith {
            param($Identity)
            $policy = [pscustomobject]@{ Identity = 'Office365 AntiPhish Default'; IsDefault = $true; EnableSpoofIntelligence = $true; EnableMailboxIntelligence = $false; EnableMailboxIntelligenceProtection = $false }
            if ($Identity) { return $policy }
            return @($policy)
        }

        # Captured directly from the mock's own $PSBoundParameters rather than
        # inspected later via -ParameterFilter: Should -Invoke's ParameterFilter
        # reconstructs bound parameters from Pester's recorded call history, which
        # was observed (live, on a real Pester 6 run) to not reliably preserve
        # which switch parameters were explicitly bound - the mock body's own
        # $PSBoundParameters at invocation time has no such ambiguity.
        $script:CapturedSetAntiPhishPolicyCalls = [System.Collections.Generic.List[object]]::new()
        Mock -CommandName Set-AntiPhishPolicy -ModuleName ExchangeOnlineControls -MockWith {
            $script:CapturedSetAntiPhishPolicyCalls.Add(@($PSBoundParameters.Keys))
        }
    }

    Context 'ExchangeOnline-AntiPhishingSpoofIntelligence (base EOP, no license gate)' {
        It 'always runs, even without any Defender for Office 365 service plan' {
            Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @([pscustomobject]@{ ServicePlans = @([pscustomobject]@{ ServicePlanName = 'EXCHANGE_S_STANDARD'; ProvisioningStatus = 'Success' }) }) }

            $result = Set-ExchangeOnline-AntiPhishingSpoofIntelligenceState -DesiredValue ([pscustomobject]@{ enableSpoofIntelligence = $true }) -CurrentValue ([pscustomobject]@{ enableSpoofIntelligence = $false })

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-AntiPhishPolicy -ModuleName ExchangeOnlineControls -Times 1
            $script:CapturedSetAntiPhishPolicyCalls.Count | Should -Be 1
            $script:CapturedSetAntiPhishPolicyCalls[0] | Should -Contain 'EnableSpoofIntelligence'
            $script:CapturedSetAntiPhishPolicyCalls[0] | Should -Not -Contain 'EnableMailboxIntelligence'
            $script:CapturedSetAntiPhishPolicyCalls[0] | Should -Not -Contain 'EnableMailboxIntelligenceProtection'
        }
    }

    Context 'ExchangeOnline-AntiPhishingMailboxIntelligence (Defender for Office 365 Plan 1/2 gated)' {
        It 'reports Skipped-LicenseInsufficient and never calls Set-AntiPhishPolicy without ATP_ENTERPRISE/THREAT_INTELLIGENCE' {
            Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @([pscustomobject]@{ ServicePlans = @([pscustomobject]@{ ServicePlanName = 'EXCHANGE_S_STANDARD'; ProvisioningStatus = 'Success' }) }) }

            (Get-ExchangeOnline-AntiPhishingMailboxIntelligenceState).Value | Should -Be $null

            $result = Set-ExchangeOnline-AntiPhishingMailboxIntelligenceState -DesiredValue ([pscustomobject]@{ enableMailboxIntelligence = $true; enableMailboxIntelligenceProtection = $true }) -CurrentValue $null
            $result.Status | Should -Be 'Skipped-LicenseInsufficient'
            $result.Message | Should -Match 'ATP_ENTERPRISE'
            # The whole point of the gate: not even a call with just the spoof-style
            # subset happens - Set-AntiPhishPolicy is not called AT ALL when the
            # gate fails, mailbox-intelligence parameters included.
            Should -Invoke -CommandName Set-AntiPhishPolicy -ModuleName ExchangeOnlineControls -Times 0
        }

        It 'proceeds and passes only the mailbox-intelligence parameters when ATP_ENTERPRISE (Defender for O365 Plan 1) is present' {
            Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @([pscustomobject]@{ ServicePlans = @([pscustomobject]@{ ServicePlanName = 'ATP_ENTERPRISE'; ProvisioningStatus = 'Success' }) }) }

            $result = Set-ExchangeOnline-AntiPhishingMailboxIntelligenceState -DesiredValue ([pscustomobject]@{ enableMailboxIntelligence = $true; enableMailboxIntelligenceProtection = $true }) -CurrentValue ([pscustomobject]@{ enableMailboxIntelligence = $false; enableMailboxIntelligenceProtection = $false })

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-AntiPhishPolicy -ModuleName ExchangeOnlineControls -Times 1
            $script:CapturedSetAntiPhishPolicyCalls.Count | Should -Be 1
            $script:CapturedSetAntiPhishPolicyCalls[0] | Should -Contain 'EnableMailboxIntelligence'
            $script:CapturedSetAntiPhishPolicyCalls[0] | Should -Contain 'EnableMailboxIntelligenceProtection'
            $script:CapturedSetAntiPhishPolicyCalls[0] | Should -Not -Contain 'EnableSpoofIntelligence'
        }

        It 'proceeds when THREAT_INTELLIGENCE (Defender for O365 Plan 2) is present instead' {
            Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @([pscustomobject]@{ ServicePlans = @([pscustomobject]@{ ServicePlanName = 'THREAT_INTELLIGENCE'; ProvisioningStatus = 'Success' }) }) }
            (Get-ExchangeOnline-AntiPhishingMailboxIntelligenceState).Value | Should -Not -Be $null
        }
    }
}

Describe 'No stale references to the old single ExchangeOnline-AntiPhishing control id remain' {

    It 'config/baseline.config.json does not reference the pre-split control id' {
        $configPath = Join-Path $PSScriptRoot '../config/baseline.config.json'
        $config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json -Depth 25
        $config.controls.id | Should -Not -Contain 'ExchangeOnline-AntiPhishing'
        $config.controls.id | Should -Contain 'ExchangeOnline-AntiPhishingSpoofIntelligence'
        $config.controls.id | Should -Contain 'ExchangeOnline-AntiPhishingMailboxIntelligence'
    }

    It 'the schema/config/README do not contain the literal string "ExchangeOnline-AntiPhishing" followed by anything other than SpoofIntelligence/MailboxIntelligence' {
        $paths = @(
            (Join-Path $PSScriptRoot '../config/baseline.config.schema.json')
            (Join-Path $PSScriptRoot '../config/baseline.config.json')
            (Join-Path $PSScriptRoot '../README.md')
        )
        foreach ($path in $paths) {
            $matches = [regex]::Matches((Get-Content -LiteralPath $path -Raw), 'ExchangeOnline-AntiPhishing(?!SpoofIntelligence|MailboxIntelligence)')
            $matches.Count | Should -Be 0 -Because "found a stale bare 'ExchangeOnline-AntiPhishing' reference in $path"
        }
    }
}
