<#
    ConditionalAccessControls.Tests.ps1

    Tests the Conditional Access controls' licensing gate, overlap detection,
    idempotent create/update, and (most importantly) the report-only guarantee,
    with every Microsoft.Graph cmdlet mocked. No live tenant connection, and
    neither Microsoft.Graph nor Microsoft.Graph.Identity.SignIns, is required.
#>

BeforeAll {
    $script:ModulesDir = Join-Path $PSScriptRoot '../modules'

    function New-FakeSku {
        param([string[]]$ServicePlanNames, [string]$ProvisioningStatus = 'Success')
        [pscustomobject]@{
            ServicePlans = @($ServicePlanNames | ForEach-Object { [pscustomobject]@{ ServicePlanName = $_; ProvisioningStatus = $ProvisioningStatus } })
        }
    }

    function New-FakeCAPolicy {
        param(
            [string]$Id = 'existing-policy-id',
            [string]$DisplayName,
            [string]$State = 'enabled',
            [string[]]$IncludeUsers = @(),
            [string[]]$IncludeRoles = @(),
            [string[]]$ExcludeGroups = @(),
            [string[]]$IncludeApplications = @(),
            [string[]]$IncludeUserActions = @(),
            [string[]]$ClientAppTypes = @(),
            [string[]]$SignInRiskLevels = @(),
            [string[]]$UserRiskLevels = @(),
            [hashtable]$IncludeGuestsOrExternalUsers = $null,
            [string]$GrantOperator = 'OR',
            [string[]]$BuiltInControls = @('mfa')
        )
        [pscustomobject]@{
            Id            = $Id
            DisplayName   = $DisplayName
            State         = $State
            Conditions    = [pscustomobject]@{
                Users          = [pscustomobject]@{
                    IncludeUsers = $IncludeUsers
                    IncludeRoles = $IncludeRoles
                    ExcludeGroups = $ExcludeGroups
                    IncludeGuestsOrExternalUsers = if ($IncludeGuestsOrExternalUsers) { [pscustomobject]$IncludeGuestsOrExternalUsers } else { [pscustomobject]@{ GuestOrExternalUserTypes = $null } }
                }
                Applications   = [pscustomobject]@{ IncludeApplications = $IncludeApplications; IncludeUserActions = $IncludeUserActions }
                ClientAppTypes = $ClientAppTypes
                SignInRiskLevels = $SignInRiskLevels
                UserRiskLevels   = $UserRiskLevels
            }
            GrantControls = [pscustomobject]@{ Operator = $GrantOperator; BuiltInControls = $BuiltInControls }
        }
    }
}

BeforeEach {
    # Re-imports both modules with -Force on every test, exactly like
    # Invoke-M365Baseline.ps1 does on every real run - this is what resets
    # BaselineCore's Get-BaselineSubscribedSkuCache and this module's own
    # CAPolicyListCache/CAEmergencyGroupIdCache/CAAdminRoleIdCache module-scoped
    # caches between tests. Without this, one test's mocked "tenant license" or
    # "existing policies" would silently leak into the next.
    Import-Module (Join-Path $script:ModulesDir 'BaselineCore.psm1') -Force
    Import-Module (Join-Path $script:ModulesDir 'ConditionalAccessControls.psm1') -Force -WarningAction SilentlyContinue

    Mock -CommandName Get-MgGroup -ModuleName ConditionalAccessControls -MockWith { $null }
    Mock -CommandName New-MgGroup -ModuleName ConditionalAccessControls -MockWith {
        param($DisplayName, [switch]$MailEnabled, $MailNickname, [switch]$SecurityEnabled)
        [pscustomobject]@{ Id = 'emergency-group-id'; DisplayName = $DisplayName }
    }
    Mock -CommandName Get-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith { @() }
    Mock -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith {
        param($BodyParameter)
        [pscustomobject]@{ Id = 'new-policy-id'; DisplayName = $BodyParameter.displayName }
    }
    Mock -CommandName Update-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith { }
    Mock -CommandName Get-MgDirectoryRoleTemplate -ModuleName ConditionalAccessControls -MockWith {
        @(
            'Global Administrator', 'Privileged Role Administrator', 'Application Administrator',
            'Cloud Application Administrator', 'Authentication Administrator', 'Privileged Authentication Administrator',
            'Security Administrator', 'Exchange Administrator', 'SharePoint Administrator', 'User Administrator',
            'Helpdesk Administrator', 'Conditional Access Administrator', 'Billing Administrator', 'Password Administrator'
        ) | ForEach-Object { [pscustomobject]@{ Id = "role-$_"; DisplayName = $_ } }
    }
    Mock -CommandName Get-MgServicePrincipal -ModuleName ConditionalAccessControls -MockWith {
        [pscustomobject]@{ Id = 'azure-mgmt-sp-id'; AppId = '797f4846-ba00-4fd7-ba43-dac1f8f63013' }
    }
}

Describe 'ConditionalAccessControls - licensing gate' {

    Context 'Entra ID Free tenant (no AAD_PREMIUM/AAD_PREMIUM_P2)' {
        BeforeEach {
            Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @(New-FakeSku -ServicePlanNames @('EXCHANGE_S_STANDARD')) }
        }

        It 'Get- reports Value = $null (Unknown), not an error, for a Tier 1 control' {
            (Get-CA-RequireMfaAllUsersState).Value | Should -Be $null
        }

        It 'Set- reports Skipped-LicenseInsufficient and never calls New-MgIdentityConditionalAccessPolicy' {
            $result = Set-CA-RequireMfaAllUsersState -DesiredValue $true -CurrentValue $null
            $result.Status | Should -Be 'Skipped-LicenseInsufficient'
            Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 0
        }

        It 'does not create the emergency-access group from a Get- call alone (Audit changes nothing)' {
            Get-CA-RequireMfaAllUsersState | Out-Null
            Should -Invoke -CommandName New-MgGroup -ModuleName ConditionalAccessControls -Times 0
        }
    }

    Context 'Entra ID P1 tenant (AAD_PREMIUM only)' {
        BeforeEach {
            Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @(New-FakeSku -ServicePlanNames @('AAD_PREMIUM')) }
        }

        It 'a Tier 1 control creates a new report-only policy' {
            $result = Set-CA-RequireMfaAllUsersState -DesiredValue $true -CurrentValue $false
            $result.Status | Should -Be 'Created'
            Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 1 -ParameterFilter {
                $BodyParameter.state -eq 'enabledForReportingButNotEnforced'
            }
        }

        It 'a Tier 2 control reports Skipped-LicenseInsufficient (P1 alone does not satisfy Tier 2)' {
            $result = Set-CA-RequireMfaSignInRiskState -DesiredValue $true -CurrentValue $null
            $result.Status | Should -Be 'Skipped-LicenseInsufficient'
            Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 0
        }
    }

    Context 'Entra ID P2 tenant (AAD_PREMIUM_P2)' {
        BeforeEach {
            Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @(New-FakeSku -ServicePlanNames @('AAD_PREMIUM_P2')) }
        }

        It 'both a Tier 1 and a Tier 2 control proceed (P2 implies P1)' {
            (Set-CA-RequireMfaAllUsersState -DesiredValue $true -CurrentValue $false).Status | Should -Be 'Created'
            (Set-CA-RequireMfaSignInRiskState -DesiredValue $true -CurrentValue $false).Status | Should -Be 'Created'
        }
    }

    It 'calls Get-MgSubscribedSku at most once per run despite many license checks across many controls' {
        Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @(New-FakeSku -ServicePlanNames @('AAD_PREMIUM_P2')) }
        Get-CA-RequireMfaAllUsersState | Out-Null
        Get-CA-RequireMfaAdminRolesState | Out-Null
        Get-CA-BlockLegacyAuthState | Out-Null
        Get-CA-RequireMfaSignInRiskState | Out-Null
        Set-CA-RequirePasswordChangeUserRiskState -DesiredValue $true -CurrentValue $false | Out-Null
        Should -Invoke -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -Times 1
    }
}

Describe 'ConditionalAccessControls - overlap detection' {

    BeforeEach {
        Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @(New-FakeSku -ServicePlanNames @('AAD_PREMIUM')) }
    }

    It 'skips creation (Skipped-PotentialOverlap) when a non-toolkit-owned policy heuristically matches, naming the conflict' {
        Mock -CommandName Get-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith {
            @(New-FakeCAPolicy -Id 'contoso-1' -DisplayName 'Contoso - Require MFA Everyone' -IncludeUsers @('All') -IncludeApplications @('All') -BuiltInControls @('mfa'))
        }
        $result = Set-CA-RequireMfaAllUsersState -DesiredValue $true -CurrentValue $false
        $result.Status | Should -Be 'Skipped-PotentialOverlap'
        $result.Message | Should -Match 'Contoso - Require MFA Everyone'
        Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 0
    }

    It 'creates the toolkit-owned policy anyway when ForceCreateDespiteOverlap is set' {
        Mock -CommandName Get-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith {
            @(New-FakeCAPolicy -Id 'contoso-1' -DisplayName 'Contoso - Require MFA Everyone' -IncludeUsers @('All') -IncludeApplications @('All') -BuiltInControls @('mfa'))
        }
        $result = Set-CA-RequireMfaAllUsersState -DesiredValue $true -CurrentValue $false -ForceCreateDespiteOverlap $true
        $result.Status | Should -Be 'Created'
        Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 1
    }

    It 'does not treat an existing toolkit-owned policy as an overlap (updates it instead of skipping)' {
        Mock -CommandName Get-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith {
            @(New-FakeCAPolicy -Id 'toolkit-1' -DisplayName '[M365 Baseline] Require MFA for all users' -State 'enabledForReportingButNotEnforced' -IncludeUsers @('All') -IncludeApplications @('All') -ExcludeGroups @('emergency-group-id') -BuiltInControls @('block'))
        }
        Mock -CommandName Get-MgGroup -ModuleName ConditionalAccessControls -MockWith { [pscustomobject]@{ Id = 'emergency-group-id'; DisplayName = 'x' } }
        $result = Set-CA-RequireMfaAllUsersState -DesiredValue $true -CurrentValue $false
        $result.Status | Should -Be 'Updated'
        Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 0
        Should -Invoke -CommandName Update-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 1
    }
}

Describe 'ConditionalAccessControls - idempotency and report-only enforcement' {

    BeforeEach {
        Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @(New-FakeSku -ServicePlanNames @('AAD_PREMIUM_P2')) }
        Mock -CommandName Get-MgGroup -ModuleName ConditionalAccessControls -MockWith { [pscustomobject]@{ Id = 'emergency-group-id'; DisplayName = 'x' } }
    }

    It 'Get- reports compliant when the existing toolkit-owned policy matches exactly' {
        Mock -CommandName Get-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith {
            @(New-FakeCAPolicy -Id 'toolkit-1' -DisplayName '[M365 Baseline] Require MFA for all users' -State 'enabledForReportingButNotEnforced' -IncludeUsers @('All') -IncludeApplications @('All') -ExcludeGroups @('emergency-group-id') -BuiltInControls @('mfa'))
        }
        (Get-CA-RequireMfaAllUsersState).Value | Should -Be $true
    }

    It 'Set- is a no-op (Success, no API call) when CurrentValue is already $true' {
        $result = Set-CA-RequireMfaAllUsersState -DesiredValue $true -CurrentValue $true
        $result.Status | Should -Be 'Success'
        $result.Message | Should -Match 'Already compliant'
        Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 0
        Should -Invoke -CommandName Update-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 0
    }

    It 'every New-MgIdentityConditionalAccessPolicy call across a Created and an Updated scenario uses state=enabledForReportingButNotEnforced and excludes the emergency-access group' {
        # Created path
        Mock -CommandName Get-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith { @() }
        Set-CA-BlockLegacyAuthState -DesiredValue $true -CurrentValue $false | Out-Null
        Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 1 -ParameterFilter {
            $BodyParameter.state -eq 'enabledForReportingButNotEnforced' -and $BodyParameter.conditions.users.excludeGroups -contains 'emergency-group-id'
        }

        # Updated path (drifted existing toolkit-owned policy)
        Mock -CommandName Get-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith {
            @(New-FakeCAPolicy -Id 'toolkit-2' -DisplayName '[M365 Baseline] Block legacy authentication' -State 'enabled' -IncludeUsers @('All') -IncludeApplications @('All') -ExcludeGroups @('emergency-group-id') -ClientAppTypes @('exchangeActiveSync', 'other') -BuiltInControls @('block'))
        }
        Set-CA-BlockLegacyAuthState -DesiredValue $true -CurrentValue $false | Out-Null
        Should -Invoke -CommandName Update-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 1 -ParameterFilter {
            $BodyParameter.state -eq 'enabledForReportingButNotEnforced'
        }
    }

    It 'never calls New-/Update-MgIdentityConditionalAccessPolicy with state set to anything other than enabledForReportingButNotEnforced, across every scenario in this Describe block' {
        Mock -CommandName Get-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -MockWith { @() }
        1..2 | ForEach-Object { Set-CA-RequireMfaAllUsersState -DesiredValue $true -CurrentValue $false | Out-Null }
        Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -ParameterFilter {
            $BodyParameter.state -ne 'enabledForReportingButNotEnforced'
        } -Times 0
        Should -Invoke -CommandName Update-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -ParameterFilter {
            $BodyParameter.state -ne 'enabledForReportingButNotEnforced'
        } -Times 0
    }
}

Describe 'ConditionalAccessControls - static guarantee: no code path ever writes state = enabled' {

    It 'the module source never assigns the literal ''enabled'' to a policy state (only ''enabledForReportingButNotEnforced'')' {
        $source = Get-Content (Join-Path $script:ModulesDir 'ConditionalAccessControls.psm1') -Raw
        # Targets an actual assignment (=  'enabled'), not a bare substring search -
        # the module's own header comment narrates the invariant in prose ("...sets
        # a CA policy's state to 'enabled'"), which a plain "'enabled'" substring
        # search would false-positive on despite it not being code. This pattern
        # also correctly does not match "state = 'enabledForReportingButNotEnforced'"
        # since the character immediately after "enabled" there is "F", not the
        # closing quote this regex requires.
        $source | Should -Not -Match "=\s*'enabled'"
        $source | Should -Match "'enabledForReportingButNotEnforced'"
    }
}

Describe 'ConditionalAccessControls - dynamic resolution (no hardcoded GUIDs)' {

    BeforeEach {
        Mock -CommandName Get-MgSubscribedSku -ModuleName BaselineCore -MockWith { @(New-FakeSku -ServicePlanNames @('AAD_PREMIUM')) }
    }

    It 'resolves admin role template ids via Get-MgDirectoryRoleTemplate rather than hardcoded GUIDs' {
        Set-CA-RequireMfaAdminRolesState -DesiredValue $true -CurrentValue $false | Out-Null
        Should -Invoke -CommandName Get-MgDirectoryRoleTemplate -ModuleName ConditionalAccessControls -Times 1
        Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 1 -ParameterFilter {
            @($BodyParameter.conditions.users.includeRoles).Count -eq 14
        }
    }

    It 'throws an actionable error when a required admin role cannot be resolved' {
        Mock -CommandName Get-MgDirectoryRoleTemplate -ModuleName ConditionalAccessControls -MockWith { @() }
        { Set-CA-RequireMfaAdminRolesState -DesiredValue $true -CurrentValue $false } | Should -Throw '*could not resolve*'
    }

    It 'verifies the Azure Management service principal via Get-MgServicePrincipal before referencing its app id' {
        Set-CA-RequireMfaAzureManagementState -DesiredValue $true -CurrentValue $false | Out-Null
        Should -Invoke -CommandName Get-MgServicePrincipal -ModuleName ConditionalAccessControls -Times 1 -ParameterFilter {
            $Filter -match '797f4846-ba00-4fd7-ba43-dac1f8f63013'
        }
    }

    It 'refuses to create the Azure Management policy if the service principal does not resolve in this tenant' {
        Mock -CommandName Get-MgServicePrincipal -ModuleName ConditionalAccessControls -MockWith { $null }
        { Set-CA-RequireMfaAzureManagementState -DesiredValue $true -CurrentValue $false } | Should -Throw '*did not resolve*'
        Should -Invoke -CommandName New-MgIdentityConditionalAccessPolicy -ModuleName ConditionalAccessControls -Times 0
    }
}
