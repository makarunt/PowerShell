<#
    ExchangeOnlineControls.Tests.ps1

    Tests representative Exchange Online controls' Get-/Set- functions with
    mocked ExchangeOnlineManagement cmdlets. No live EXO connection or the
    ExchangeOnlineManagement module itself is required.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
    Import-Module (Join-Path $PSScriptRoot '../modules/ExchangeOnlineControls.psm1') -Force
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
