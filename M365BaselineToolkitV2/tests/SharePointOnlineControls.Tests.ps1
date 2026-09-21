<#
    SharePointOnlineControls.Tests.ps1

    Tests the four v2 SharePoint Online controls with mocked Get-SPOTenant/
    Set-SPOTenant cmdlets. No live SPO connection or the
    Microsoft.Online.SharePoint.PowerShell module itself is required - Pester's
    Mock synthesizes the mocked commands. (v1 shipped this module without a
    dedicated test file; this v2-only file covers the four new controls it
    gained, not a retroactive full-module test pass on the pre-existing six.)
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '../modules/SharePointOnlineControls.psm1') -Force
}

Describe 'SharePointOnline-AzureADB2BIntegration' {

    Context 'Get-SharePointOnline-AzureADB2BIntegrationState' {
        It 'reads EnableAzureADB2BIntegration from the tenant' {
            Mock -CommandName Get-SPOTenant -ModuleName SharePointOnlineControls -MockWith {
                [pscustomobject]@{ EnableAzureADB2BIntegration = $true }
            }
            (Get-SharePointOnline-AzureADB2BIntegrationState).Value | Should -Be $true
        }
    }

    Context 'Set-SharePointOnline-AzureADB2BIntegrationState' {
        It 'calls Set-SPOTenant -EnableAzureADB2BIntegration when non-compliant' {
            Mock -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -MockWith { }
            $result = Set-SharePointOnline-AzureADB2BIntegrationState -DesiredValue $true -CurrentValue $false
            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -Times 1 -ParameterFilter {
                $EnableAzureADB2BIntegration -eq $true
            }
        }

        It 'is a no-op when already compliant' {
            Mock -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -MockWith { }
            $result = Set-SharePointOnline-AzureADB2BIntegrationState -DesiredValue $true -CurrentValue $true
            $result.Message | Should -Match 'Already compliant'
            Should -Invoke -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -Times 0
        }
    }
}

Describe 'SharePointOnline-PreventGuestResharing' {

    Context 'Get-SharePointOnline-PreventGuestResharingState' {
        It 'reads PreventExternalUsersFromResharing from the tenant' {
            Mock -CommandName Get-SPOTenant -ModuleName SharePointOnlineControls -MockWith {
                [pscustomobject]@{ PreventExternalUsersFromResharing = $false }
            }
            (Get-SharePointOnline-PreventGuestResharingState).Value | Should -Be $false
        }
    }

    Context 'Set-SharePointOnline-PreventGuestResharingState' {
        It 'calls Set-SPOTenant -PreventExternalUsersFromResharing when non-compliant' {
            Mock -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -MockWith { }
            $result = Set-SharePointOnline-PreventGuestResharingState -DesiredValue $true -CurrentValue $false
            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -Times 1 -ParameterFilter {
                $PreventExternalUsersFromResharing -eq $true
            }
        }
    }
}

Describe 'SharePointOnline-GuestAccessExpiration' {

    Context 'Get-SharePointOnline-GuestAccessExpirationState' {
        It 'reads ExternalUserExpirationRequired/ExternalUserExpireInDays from the tenant' {
            Mock -CommandName Get-SPOTenant -ModuleName SharePointOnlineControls -MockWith {
                [pscustomobject]@{ ExternalUserExpirationRequired = $true; ExternalUserExpireInDays = 30 }
            }
            $result = Get-SharePointOnline-GuestAccessExpirationState
            $result.Value.externalUserExpirationRequired | Should -Be $true
            $result.Value.externalUserExpireInDays | Should -Be 30
        }
    }

    Context 'Set-SharePointOnline-GuestAccessExpirationState' {
        It 'calls Set-SPOTenant with both fields when non-compliant' {
            Mock -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -MockWith { }
            $current = [pscustomobject]@{ externalUserExpirationRequired = $false; externalUserExpireInDays = 60 }
            $desired = [pscustomobject]@{ externalUserExpirationRequired = $true; externalUserExpireInDays = 30 }
            $result = Set-SharePointOnline-GuestAccessExpirationState -DesiredValue $desired -CurrentValue $current
            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -Times 1 -ParameterFilter {
                $ExternalUserExpirationRequired -eq $true -and $ExternalUserExpireInDays -eq 30
            }
        }
    }

    It 'is a distinct control id from SharePointOnline-AnonymousLinkExpiration' {
        (Get-Command Get-SharePointOnline-GuestAccessExpirationState).Name | Should -Not -Be (Get-Command Get-SharePointOnline-AnonymousLinkExpirationState).Name
    }
}

Describe 'SharePointOnline-GuestReauthentication' {

    Context 'Get-SharePointOnline-GuestReauthenticationState' {
        It 'reads EmailAttestationRequired/EmailAttestationReAuthDays from the tenant' {
            Mock -CommandName Get-SPOTenant -ModuleName SharePointOnlineControls -MockWith {
                [pscustomobject]@{ EmailAttestationRequired = $true; EmailAttestationReAuthDays = 15 }
            }
            $result = Get-SharePointOnline-GuestReauthenticationState
            $result.Value.emailAttestationRequired | Should -Be $true
            $result.Value.emailAttestationReAuthDays | Should -Be 15
        }
    }

    Context 'Set-SharePointOnline-GuestReauthenticationState' {
        It 'calls Set-SPOTenant with both fields when non-compliant' {
            Mock -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -MockWith { }
            $current = [pscustomobject]@{ emailAttestationRequired = $false; emailAttestationReAuthDays = 45 }
            $desired = [pscustomobject]@{ emailAttestationRequired = $true; emailAttestationReAuthDays = 15 }
            $result = Set-SharePointOnline-GuestReauthenticationState -DesiredValue $desired -CurrentValue $current
            $result.Status | Should -Be 'Success'
            Should -Invoke -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -Times 1 -ParameterFilter {
                $EmailAttestationRequired -eq $true -and $EmailAttestationReAuthDays -eq 15
            }
        }
    }
}

Describe 'Read-back-and-classify integration: B2B deprecation and guest-resharing propagation delay' {

    BeforeAll {
        Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
    }

    It 'a control flagged mechanismPossiblyDeprecated reports MechanismPossiblyDeprecated when Set- succeeds but the read-back never matches' {
        $config = [pscustomobject]@{
            controls = @(
                [pscustomobject]@{
                    id = 'SharePointOnline-AzureADB2BIntegration'; workload = 'SharePointOnline'; enabled = $true
                    automatable = $true; desiredValue = $true; description = 'test'
                    mechanismPossiblyDeprecated = $true
                }
            )
        }
        Mock -CommandName Get-SPOTenant -ModuleName SharePointOnlineControls -MockWith { [pscustomobject]@{ EnableAzureADB2BIntegration = $false } }
        Mock -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -MockWith { }

        $catalog = Get-BaselineControlCatalog -Config $config -AvailableFunctions @(
            'Get-SharePointOnline-AzureADB2BIntegrationState', 'Set-SharePointOnline-AzureADB2BIntegrationState'
        )
        $catalog[0].MechanismPossiblyDeprecated | Should -Be $true

        $audit = Invoke-BaselineControlAudit -Catalog $catalog
        $changeLogPath = [System.IO.Path]::GetTempFileName()
        try {
            $applyResults = Invoke-BaselineControlApply -Catalog $catalog -AuditResults $audit -ChangeLogPath $changeLogPath -ApplyOutcomeRetryDelayMilliseconds 0
            $applyResults[0].Status | Should -Be 'MechanismPossiblyDeprecated'
        }
        finally {
            Remove-Item $changeLogPath -ErrorAction SilentlyContinue
        }
    }

    It 'a control NOT flagged mechanismPossiblyDeprecated reports Applied-PendingConfirmation for the same unconfirmed-write pattern' {
        $config = [pscustomobject]@{
            controls = @(
                [pscustomobject]@{
                    id = 'SharePointOnline-PreventGuestResharing'; workload = 'SharePointOnline'; enabled = $true
                    automatable = $true; desiredValue = $true; description = 'test'
                }
            )
        }
        Mock -CommandName Get-SPOTenant -ModuleName SharePointOnlineControls -MockWith { [pscustomobject]@{ PreventExternalUsersFromResharing = $false } }
        Mock -CommandName Set-SPOTenant -ModuleName SharePointOnlineControls -MockWith { }

        $catalog = Get-BaselineControlCatalog -Config $config -AvailableFunctions @(
            'Get-SharePointOnline-PreventGuestResharingState', 'Set-SharePointOnline-PreventGuestResharingState'
        )
        $catalog[0].MechanismPossiblyDeprecated | Should -Be $false

        $audit = Invoke-BaselineControlAudit -Catalog $catalog
        $changeLogPath = [System.IO.Path]::GetTempFileName()
        try {
            $applyResults = Invoke-BaselineControlApply -Catalog $catalog -AuditResults $audit -ChangeLogPath $changeLogPath -ApplyOutcomeRetryDelayMilliseconds 0
            $applyResults[0].Status | Should -Be 'Applied-PendingConfirmation'
        }
        finally {
            Remove-Item $changeLogPath -ErrorAction SilentlyContinue
        }
    }
}
