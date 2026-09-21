<#
    EntraIdControls.Tests.ps1

    Tests representative EntraID controls' Get-/Set- functions with mocked
    Microsoft.Graph cmdlets. No live Graph connection or the Microsoft.Graph
    module itself is required - Pester's Mock synthesizes the mocked commands.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '../modules/EntraIdControls.psm1') -Force
}

Describe 'EntraID-GuestInviteRestriction' {

    Context 'Get-EntraID-GuestInviteRestrictionState' {
        It 'reads AllowInvitesFrom from the authorization policy' {
            Mock -CommandName Get-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -MockWith {
                [pscustomobject]@{ Id = 'authorizationPolicy'; AllowInvitesFrom = 'everyone' }
            }
            $result = Get-EntraID-GuestInviteRestrictionState
            $result.Value | Should -Be 'everyone'
        }
    }

    Context 'Set-EntraID-GuestInviteRestrictionState' {
        It 'is a no-op when already compliant' {
            Mock -CommandName Update-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -MockWith { }
            $result = Set-EntraID-GuestInviteRestrictionState -DesiredValue 'adminsAndGuestInviters' -CurrentValue 'adminsAndGuestInviters'
            $result.Status | Should -Be 'Success'
            $result.Message | Should -Match 'Already compliant'
            Should -Invoke -CommandName Update-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -Times 0
        }

        It 'calls Update-MgPolicyAuthorizationPolicy with the desired value when non-compliant' {
            Mock -CommandName Get-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -MockWith {
                [pscustomobject]@{ Id = 'authorizationPolicy'; AllowInvitesFrom = 'everyone' }
            }
            Mock -CommandName Update-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -MockWith { }

            $result = Set-EntraID-GuestInviteRestrictionState -DesiredValue 'adminsAndGuestInviters' -CurrentValue 'everyone'

            $result.Status | Should -Be 'Success'
            $result.AppliedValue | Should -Be 'adminsAndGuestInviters'
            Should -Invoke -CommandName Update-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -Times 1 -ParameterFilter {
                $AllowInvitesFrom -eq 'adminsAndGuestInviters' -and $AuthorizationPolicyId -eq 'authorizationPolicy'
            }
        }
    }
}

Describe 'EntraID-GlobalAdminCount (audit-only)' {

    Context 'Get-EntraID-GlobalAdminCountState' {
        It 'counts Global Administrator role members' {
            Mock -CommandName Get-MgDirectoryRole -ModuleName EntraIdControls -MockWith {
                [pscustomobject]@{ Id = 'role-guid'; DisplayName = 'Global Administrator' }
            }
            Mock -CommandName Get-MgDirectoryRoleMember -ModuleName EntraIdControls -MockWith {
                @([pscustomobject]@{ Id = 'a' }, [pscustomobject]@{ Id = 'b' }, [pscustomobject]@{ Id = 'c' })
            }

            $result = Get-EntraID-GlobalAdminCountState
            $result.Value | Should -Be 3
        }

        It 'throws a clear error when the Global Administrator role is not found' {
            Mock -CommandName Get-MgDirectoryRole -ModuleName EntraIdControls -MockWith { $null }
            { Get-EntraID-GlobalAdminCountState } | Should -Throw '*not activated*'
        }
    }

    Context 'Set-EntraID-GlobalAdminCountState' {
        It 'never calls any mutating Graph cmdlet and always returns Skipped-Manual' {
            $result = Set-EntraID-GlobalAdminCountState -DesiredValue @{ min = 2; max = 4 } -CurrentValue 6
            $result.Status | Should -Be 'Skipped-Manual'
            $result.Message | Should -Match 'Entra admin center'
        }
    }
}

Describe 'EntraID-AuthMethodsHardening' {

    Context 'Get-EntraID-AuthMethodsHardeningState' {
        It 'assembles authenticator/sms/voice/systemCredentialPreferences state' {
            Mock -CommandName Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -MockWith {
                param($AuthenticationMethodConfigurationId)
                switch ($AuthenticationMethodConfigurationId) {
                    'MicrosoftAuthenticator' { [pscustomobject]@{ State = 'enabled' } }
                    'Sms' { [pscustomobject]@{ State = 'disabled' } }
                    'Voice' { [pscustomobject]@{ State = 'disabled' } }
                }
            }
            Mock -CommandName Get-MgPolicyAuthenticationMethodPolicy -ModuleName EntraIdControls -MockWith {
                [pscustomobject]@{ SystemCredentialPreferences = [pscustomobject]@{ State = 'enabled' } }
            }

            $result = Get-EntraID-AuthMethodsHardeningState
            $result.Value.authenticatorEnabled | Should -Be $true
            $result.Value.smsEnabled | Should -Be $false
            $result.Value.voiceEnabled | Should -Be $false
            $result.Value.systemCredentialPreferences.state | Should -Be 'enabled'
            $result.Detail | Should -Match 'September 2026'
        }
    }

    Context 'Set-EntraID-AuthMethodsHardeningState' {
        It 'updates all three method configurations AND systemCredentialPreferences when non-compliant' {
            Mock -CommandName Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -MockWith { }
            Mock -CommandName Update-MgPolicyAuthenticationMethodPolicy -ModuleName EntraIdControls -MockWith { }

            $current = [pscustomobject]@{ authenticatorEnabled = $false; smsEnabled = $true; voiceEnabled = $true; systemCredentialPreferences = @{ state = 'disabled' } }
            $desired = [pscustomobject]@{ authenticatorEnabled = $true; smsEnabled = $false; voiceEnabled = $false; systemCredentialPreferences = @{ state = 'enabled' } }

            $result = Set-EntraID-AuthMethodsHardeningState -DesiredValue $desired -CurrentValue $current

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -Times 3
            Should -Invoke -CommandName Update-MgPolicyAuthenticationMethodPolicy -ModuleName EntraIdControls -Times 1 -ParameterFilter {
                $BodyParameter.systemCredentialPreferences.state -eq 'enabled'
            }
        }

        It 'is a no-op when already compliant (including systemCredentialPreferences)' {
            Mock -CommandName Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -MockWith { }
            Mock -CommandName Update-MgPolicyAuthenticationMethodPolicy -ModuleName EntraIdControls -MockWith { }

            $same = [pscustomobject]@{ authenticatorEnabled = $true; smsEnabled = $false; voiceEnabled = $false; systemCredentialPreferences = @{ state = 'enabled' } }
            $result = Set-EntraID-AuthMethodsHardeningState -DesiredValue $same -CurrentValue $same

            $result.Message | Should -Match 'Already compliant'
            Should -Invoke -CommandName Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -Times 0
            Should -Invoke -CommandName Update-MgPolicyAuthenticationMethodPolicy -ModuleName EntraIdControls -Times 0
        }
    }
}

Describe 'EntraID-AdminConsentWorkflow' {

    Context 'Get-EntraID-AdminConsentWorkflowState' {
        It 'reads the admin consent request policy, including reviewers' {
            Mock -CommandName Get-MgPolicyAdminConsentRequestPolicy -ModuleName EntraIdControls -MockWith {
                [pscustomobject]@{
                    IsEnabled = $true; NotifyReviewers = $true; RemindersEnabled = $true; RequestDurationInDays = 30
                    Reviewers = @([pscustomobject]@{ Query = '/v1.0/users/abc'; QueryType = 'MicrosoftGraph'; QueryRoot = $null })
                }
            }
            $result = Get-EntraID-AdminConsentWorkflowState
            $result.Value.isEnabled | Should -Be $true
            $result.Value.requestDurationInDays | Should -Be 30
            $result.Value.reviewers.Count | Should -Be 1
            $result.Value.reviewers[0].query | Should -Be '/v1.0/users/abc'
        }
    }

    Context 'Set-EntraID-AdminConsentWorkflowState' {
        It 'throws a clear error when desiredValue.reviewers is empty' {
            $desired = [pscustomobject]@{ isEnabled = $true; notifyReviewers = $true; remindersEnabled = $true; requestDurationInDays = 30; reviewers = @() }
            $current = [pscustomobject]@{ isEnabled = $false; notifyReviewers = $false; remindersEnabled = $false; requestDurationInDays = 7; reviewers = @() }
            { Set-EntraID-AdminConsentWorkflowState -DesiredValue $desired -CurrentValue $current } | Should -Throw '*requires at least one entry*'
        }

        It 'calls Update-MgPolicyAdminConsentRequestPolicy with all fields including reviewers when populated' {
            Mock -CommandName Update-MgPolicyAdminConsentRequestPolicy -ModuleName EntraIdControls -MockWith { }
            $current = [pscustomobject]@{ isEnabled = $false; notifyReviewers = $false; remindersEnabled = $false; requestDurationInDays = 7; reviewers = @() }
            $desired = [pscustomobject]@{
                isEnabled = $true; notifyReviewers = $true; remindersEnabled = $true; requestDurationInDays = 30
                reviewers = @(@{ query = '/v1.0/users/00000000-0000-0000-0000-000000000000'; queryType = 'MicrosoftGraph' })
            }

            $result = Set-EntraID-AdminConsentWorkflowState -DesiredValue $desired -CurrentValue $current

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Update-MgPolicyAdminConsentRequestPolicy -ModuleName EntraIdControls -Times 1 -ParameterFilter {
                $RequestDurationInDays -eq 30 -and $Reviewers.Count -eq 1
            }
        }

        It 'is a no-op when already compliant' {
            Mock -CommandName Update-MgPolicyAdminConsentRequestPolicy -ModuleName EntraIdControls -MockWith { }
            $same = [pscustomobject]@{
                isEnabled = $true; notifyReviewers = $true; remindersEnabled = $true; requestDurationInDays = 30
                reviewers = @(@{ query = '/v1.0/users/x'; queryType = 'MicrosoftGraph' })
            }
            $result = Set-EntraID-AdminConsentWorkflowState -DesiredValue $same -CurrentValue $same
            $result.Message | Should -Match 'Already compliant'
            Should -Invoke -CommandName Update-MgPolicyAdminConsentRequestPolicy -ModuleName EntraIdControls -Times 0
        }
    }
}

Describe 'EntraID-GaNotLocalAdminOnJoin (audit-only)' {

    Context 'Get-EntraID-GaNotLocalAdminOnJoinState' {
        It 'always reports Value = $null with a Detail explaining why' {
            $result = Get-EntraID-GaNotLocalAdminOnJoinState
            $result.Value | Should -Be $null
            $result.Detail | Should -Match 'Preview'
        }
    }

    Context 'Set-EntraID-GaNotLocalAdminOnJoinState' {
        It 'never calls any mutating Graph cmdlet and always returns Skipped-Manual' {
            $result = Set-EntraID-GaNotLocalAdminOnJoinState -DesiredValue $true -CurrentValue $null
            $result.Status | Should -Be 'Skipped-Manual'
            $result.Message | Should -Match 'Device settings'
        }
    }
}

Describe 'EntraID-BlockSelfServiceAppCreation / EntraID-BlockSelfServiceSecurityGroupCreation overlap (v2 deviation)' {

    It 'BlockSelfServiceAppCreation does NOT gain an allowedToCreateSecurityGroups field (kept as a separate control)' {
        Mock -CommandName Get-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -MockWith {
            [pscustomobject]@{ DefaultUserRolePermissions = [pscustomobject]@{ AllowedToCreateApps = $false; AllowedToCreateTenants = $false } }
        }
        $result = Get-EntraID-BlockSelfServiceAppCreationState
        $result.Value.PSObject.Properties['allowedToCreateSecurityGroups'] | Should -Be $null
    }

    It 'BlockSelfServiceSecurityGroupCreation still independently owns allowedToCreateSecurityGroups' {
        Mock -CommandName Update-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -MockWith { }
        $result = Set-EntraID-BlockSelfServiceSecurityGroupCreationState -DesiredValue $false -CurrentValue $true
        $result.Status | Should -Be 'Success'
        Should -Invoke -CommandName Update-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -Times 1 -ParameterFilter {
            $BodyParameter.defaultUserRolePermissions.ContainsKey('allowedToCreateSecurityGroups') -and
            -not $BodyParameter.defaultUserRolePermissions.ContainsKey('allowedToCreateApps')
        }
    }
}
