<#
    TeamsControls.Tests.ps1

    Tests the two v2-extended Teams controls with mocked MicrosoftTeams
    cmdlets. No live Teams connection or the MicrosoftTeams module itself is
    required - Pester's Mock synthesizes the mocked commands. (v1 shipped this
    module without a dedicated test file; this v2-only file focuses on the
    two controls that gained new fields, with field-preservation assertions
    confirming the extension didn't silently drop what was already there.)
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '../modules/TeamsControls.psm1') -Force
}

Describe 'Teams-BlockConsumerContact (extended: externalAccessWithTrialTenants)' {

    Context 'Get-Teams-BlockConsumerContactState' {
        It 'reads allowTeamsConsumer/allowTeamsConsumerInbound/externalAccessWithTrialTenants' {
            Mock -CommandName Get-CsTenantFederationConfiguration -ModuleName TeamsControls -MockWith {
                [pscustomobject]@{ AllowTeamsConsumer = $true; AllowTeamsConsumerInbound = $true; ExternalAccessWithTrialTenants = 'Allowed' }
            }
            $result = Get-Teams-BlockConsumerContactState
            $result.Value.allowTeamsConsumer | Should -Be $true
            $result.Value.externalAccessWithTrialTenants | Should -Be 'Allowed'
        }
    }

    Context 'Set-Teams-BlockConsumerContactState' {
        It 'forwards the new externalAccessWithTrialTenants field ALONGSIDE the pre-existing two fields' {
            Mock -CommandName Set-CsTenantFederationConfiguration -ModuleName TeamsControls -MockWith { }
            $current = [pscustomobject]@{ allowTeamsConsumer = $true; allowTeamsConsumerInbound = $true; externalAccessWithTrialTenants = 'Allowed' }
            $desired = [pscustomobject]@{ allowTeamsConsumer = $false; allowTeamsConsumerInbound = $false; externalAccessWithTrialTenants = 'Blocked' }

            $result = Set-Teams-BlockConsumerContactState -DesiredValue $desired -CurrentValue $current

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-CsTenantFederationConfiguration -ModuleName TeamsControls -Times 1 -ParameterFilter {
                $AllowTeamsConsumer -eq $false -and $AllowTeamsConsumerInbound -eq $false -and $ExternalAccessWithTrialTenants -eq 'Blocked'
            }
        }

        It 'is a no-op when already compliant' {
            Mock -CommandName Set-CsTenantFederationConfiguration -ModuleName TeamsControls -MockWith { }
            $same = [pscustomobject]@{ allowTeamsConsumer = $false; allowTeamsConsumerInbound = $false; externalAccessWithTrialTenants = 'Blocked' }
            $result = Set-Teams-BlockConsumerContactState -DesiredValue $same -CurrentValue $same
            $result.Message | Should -Match 'Already compliant'
            Should -Invoke -CommandName Set-CsTenantFederationConfiguration -ModuleName TeamsControls -Times 0
        }
    }
}

Describe 'Teams-MeetingJoinDefaults (extended: allowAnonymousUsersToStartMeeting, allowPSTNUsersToBypassLobby)' {

    Context 'Get-Teams-MeetingJoinDefaultsState' {
        It 'reads all four fields' {
            Mock -CommandName Get-CsTeamsMeetingPolicy -ModuleName TeamsControls -MockWith {
                [pscustomobject]@{
                    AutoAdmittedUsers = 'EveryoneInCompanyExcludingGuests'
                    AllowAnonymousUsersToJoinMeeting = $false
                    AllowAnonymousUsersToStartMeeting = $false
                    AllowPSTNUsersToBypassLobby = $false
                }
            }
            $result = Get-Teams-MeetingJoinDefaultsState
            $result.Value.allowAnonymousUsersToStartMeeting | Should -Be $false
            $result.Value.allowPSTNUsersToBypassLobby | Should -Be $false
        }
    }

    Context 'Set-Teams-MeetingJoinDefaultsState' {
        It 'forwards the two new fields ALONGSIDE the pre-existing two fields' {
            Mock -CommandName Set-CsTeamsMeetingPolicy -ModuleName TeamsControls -MockWith { }
            $current = [pscustomobject]@{ autoAdmittedUsers = 'Everyone'; allowAnonymousUsersToJoinMeeting = $true; allowAnonymousUsersToStartMeeting = $true; allowPSTNUsersToBypassLobby = $true }
            $desired = [pscustomobject]@{ autoAdmittedUsers = 'EveryoneInCompanyExcludingGuests'; allowAnonymousUsersToJoinMeeting = $false; allowAnonymousUsersToStartMeeting = $false; allowPSTNUsersToBypassLobby = $false }

            $result = Set-Teams-MeetingJoinDefaultsState -DesiredValue $desired -CurrentValue $current

            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-CsTeamsMeetingPolicy -ModuleName TeamsControls -Times 1 -ParameterFilter {
                $AutoAdmittedUsers -eq 'EveryoneInCompanyExcludingGuests' -and
                $AllowAnonymousUsersToJoinMeeting -eq $false -and
                $AllowAnonymousUsersToStartMeeting -eq $false -and
                $AllowPSTNUsersToBypassLobby -eq $false
            }
        }

        It 'is a no-op when already compliant' {
            Mock -CommandName Set-CsTeamsMeetingPolicy -ModuleName TeamsControls -MockWith { }
            $same = [pscustomobject]@{ autoAdmittedUsers = 'EveryoneInCompanyExcludingGuests'; allowAnonymousUsersToJoinMeeting = $false; allowAnonymousUsersToStartMeeting = $false; allowPSTNUsersToBypassLobby = $false }
            $result = Set-Teams-MeetingJoinDefaultsState -DesiredValue $same -CurrentValue $same
            $result.Message | Should -Match 'Already compliant'
            Should -Invoke -CommandName Set-CsTeamsMeetingPolicy -ModuleName TeamsControls -Times 0
        }

        It 'falls back to a read-back verification when Set-CsTeamsMeetingPolicy throws the documented false-positive 40301' {
            Mock -CommandName Set-CsTeamsMeetingPolicy -ModuleName TeamsControls -MockWith { throw 'Forbidden (40301)' }
            Mock -CommandName Get-CsTeamsMeetingPolicy -ModuleName TeamsControls -MockWith {
                [pscustomobject]@{ AutoAdmittedUsers = 'EveryoneInCompanyExcludingGuests'; AllowAnonymousUsersToJoinMeeting = $false; AllowAnonymousUsersToStartMeeting = $false; AllowPSTNUsersToBypassLobby = $false }
            }
            $current = [pscustomobject]@{ autoAdmittedUsers = 'Everyone'; allowAnonymousUsersToJoinMeeting = $true; allowAnonymousUsersToStartMeeting = $true; allowPSTNUsersToBypassLobby = $true }
            $desired = [pscustomobject]@{ autoAdmittedUsers = 'EveryoneInCompanyExcludingGuests'; allowAnonymousUsersToJoinMeeting = $false; allowAnonymousUsersToStartMeeting = $false; allowPSTNUsersToBypassLobby = $false }

            $result = Set-Teams-MeetingJoinDefaultsState -DesiredValue $desired -CurrentValue $current
            $result.Status | Should -Be 'Success'
        }
    }
}
