<#
    EntraIdControls.Tests.ps1

    Tests representative EntraID controls' Get-/Set- functions with mocked
    Microsoft.Graph cmdlets. No live Graph connection or the Microsoft.Graph
    module itself is required - Pester's Mock synthesizes the mocked commands.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '../modules/EntraIdControls.psm1') -Force

    # Answers "was this field actually sent" rather than "what's its value" -
    # used to prove a Set- function's -BodyParameter is a single-field PATCH
    # that doesn't also touch a sibling field. A plain hashtable answers this
    # with .ContainsKey() (NOT .Contains(): the real Microsoft.Graph SDK's
    # AdditionalProperties is a generic Dictionary<string,object>, whose
    # PUBLIC (non-explicit) .Contains(item) overload is inherited from
    # ICollection<KeyValuePair<TKey,TValue>> and takes a KeyValuePair, not a
    # bare key - calling it with one string argument throws "Cannot find an
    # overload for 'Contains' and the argument count: 1". ContainsKey(key) is
    # the one method both Hashtable and Dictionary<TKey,TValue> expose
    # publicly with the same single-key-argument signature); a real typed SDK
    # object has every property slot present regardless, so a field that was
    # never set stays at its type's default ($null for every nullable
    # property these Graph models use) - "not sent" there means "still null",
    # which also matches how these SDK types serialize to the wire (a null
    # property is omitted from the JSON body, not sent as an explicit null).
    function Test-BaselineTestGraphBodyHasField {
        param([Parameter(Mandatory)][object]$Body, [Parameter(Mandatory)][string]$Name)
        if ($Body -is [System.Collections.IDictionary]) { return $Body.ContainsKey($Name) }
        $prop = $Body.PSObject.Properties[$Name]
        if (-not $prop) {
            $additional = $Body.PSObject.Properties['AdditionalProperties']
            if ($additional -and $additional.Value -is [System.Collections.IDictionary]) { return $additional.Value.ContainsKey($Name) }
            return $false
        }
        return $null -ne $prop.Value
    }
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
            # authorizationPolicy is a singleton (see Set-EntraID-GuestInviteRestrictionState's
            # own comment) - the real call is Update-MgPolicyAuthorizationPolicy -BodyParameter
            # @{ allowInvitesFrom = ... } -ErrorAction Stop, with no -AuthorizationPolicyId or
            # -AllowInvitesFrom parameter of its own. The value must be checked as a NESTED
            # property of -BodyParameter, not as a top-level bound parameter - dot-access on
            # $BodyParameter works whether it's still the plain hashtable Set- built (as it is
            # when the real Microsoft.Graph module isn't installed, so Pester's mock proxy has
            # no typed parameter to coerce it against) or has been coerced into the real
            # Microsoft.Graph.PowerShell.Models.MicrosoftGraphAuthorizationPolicy type (as it is
            # when that module IS installed, since Pester's mock proxy then inherits the real
            # cmdlet's typed -BodyParameter and PowerShell coerces the hashtable accordingly).
            Should -Invoke -CommandName Update-MgPolicyAuthorizationPolicy -ModuleName EntraIdControls -Times 1 -ParameterFilter {
                $BodyParameter.AllowInvitesFrom -eq 'adminsAndGuestInviters'
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

Describe 'EntraID-AuthMethodsHardening (audit-only; not automatable by design)' {

    Context 'Get-EntraID-AuthMethodsHardeningState' {
        It 'assembles authenticator/sms/voice state' {
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
        It 'never calls any mutating Graph cmdlet and always returns Skipped-Manual' {
            # Deliberately not automatable: disabling SMS/Voice tenant-wide
            # risks locking out an admin or user who still relies on one of
            # them to sign in - that call needs a human, not an unattended
            # script. Same pattern as EntraID-GaNotLocalAdminOnJoin below.
            $desired = [pscustomobject]@{ authenticatorEnabled = $true; smsEnabled = $false; voiceEnabled = $false }
            $current = [pscustomobject]@{ authenticatorEnabled = $false; smsEnabled = $true; voiceEnabled = $true }

            $result = Set-EntraID-AuthMethodsHardeningState -DesiredValue $desired -CurrentValue $current

            $result.Status | Should -Be 'Skipped-Manual'
            $result.Message | Should -Match 'Entra admin center'
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
            (Test-BaselineTestGraphBodyHasField -Body $BodyParameter.defaultUserRolePermissions -Name 'allowedToCreateSecurityGroups') -and
            -not (Test-BaselineTestGraphBodyHasField -Body $BodyParameter.defaultUserRolePermissions -Name 'allowedToCreateApps')
        }
    }
}
