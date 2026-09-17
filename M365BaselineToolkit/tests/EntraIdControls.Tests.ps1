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
        It 'assembles authenticator/sms/voice state from three method configurations' {
            Mock -CommandName Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -MockWith {
                param($AuthenticationMethodConfigurationId)
                switch ($AuthenticationMethodConfigurationId) {
                    'MicrosoftAuthenticator' { [pscustomobject]@{ State = 'enabled' } }
                    'Sms' { [pscustomobject]@{ State = 'disabled' } }
                    'Voice' { [pscustomobject]@{ State = 'disabled' } }
                }
            }

            $result = Get-EntraID-AuthMethodsHardeningState
            $result.Value.authenticatorEnabled | Should -Be $true
            $result.Value.smsEnabled | Should -Be $false
            $result.Value.voiceEnabled | Should -Be $false
        }
    }

    Context 'Set-EntraID-AuthMethodsHardeningState' {
        It 'updates all three method configurations when non-compliant' {
            Mock -CommandName Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -MockWith { }

            $current = [pscustomobject]@{ authenticatorEnabled = $false; smsEnabled = $true; voiceEnabled = $true }
            $desired = [pscustomobject]@{ authenticatorEnabled = $true; smsEnabled = $false; voiceEnabled = $false }

            $result = Set-EntraID-AuthMethodsHardeningState -DesiredValue $desired -CurrentValue $current

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -Times 3
        }

        It 'is a no-op when already compliant' {
            Mock -CommandName Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -MockWith { }

            $same = [pscustomobject]@{ authenticatorEnabled = $true; smsEnabled = $false; voiceEnabled = $false }
            $result = Set-EntraID-AuthMethodsHardeningState -DesiredValue $same -CurrentValue $same

            $result.Message | Should -Match 'Already compliant'
            Should -Invoke -CommandName Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -ModuleName EntraIdControls -Times 0
        }
    }
}
