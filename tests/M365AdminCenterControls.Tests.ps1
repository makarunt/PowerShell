<#
    M365AdminCenterControls.Tests.ps1

    Tests the M365 Admin Center org-settings workload module. Currently a
    single audit-only control with no PowerShell/Graph API - these tests
    confirm the audit-only contract (Skipped-Manual, no mutating calls, a
    Detail/manualInstructions pointer for the report) rather than mocking any
    backend cmdlet, since none exists for this control.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '../modules/M365AdminCenterControls.psm1') -Force
}

Describe 'M365AdminCenter-SwayExternalSharing (audit-only)' {

    Context 'Get-M365AdminCenter-SwayExternalSharingState' {
        It 'always reports Value = $null with a Detail explaining why' {
            $result = Get-M365AdminCenter-SwayExternalSharingState
            $result.Value | Should -Be $null
            $result.Detail | Should -Match 'No PowerShell or Microsoft Graph API'
        }
    }

    Context 'Set-M365AdminCenter-SwayExternalSharingState' {
        It 'always returns Skipped-Manual with the exact GUI path' {
            $result = Set-M365AdminCenter-SwayExternalSharingState -DesiredValue $false -CurrentValue $null
            $result.Status | Should -Be 'Skipped-Manual'
            $result.Message | Should -Match 'Org settings'
            $result.Message | Should -Match 'Sway'
        }
    }
}

Describe 'M365AdminCenter workload wiring' {
    It 'Get-BaselineControlCatalog resolves M365AdminCenter-SwayExternalSharing to a Graph connection' {
        $config = [pscustomobject]@{
            controls = @(
                [pscustomobject]@{
                    id = 'M365AdminCenter-SwayExternalSharing'; workload = 'M365AdminCenter'; enabled = $true
                    automatable = $false; desiredValue = $false; description = 'test'
                    manualInstructions = 'Change it by hand.'
                }
            )
        }
        $catalog = Get-BaselineControlCatalog -Config $config -AvailableFunctions @(
            'Get-M365AdminCenter-SwayExternalSharingState', 'Set-M365AdminCenter-SwayExternalSharingState'
        )
        $catalog.Count | Should -Be 1
        $catalog[0].Connection | Should -Be 'Graph'
        $catalog[0].Automatable | Should -Be $false
    }
}
