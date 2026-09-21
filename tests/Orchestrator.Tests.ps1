<#
    Orchestrator.Tests.ps1

    Tests BaselineCore's compliance-diffing logic and the Apply/Restore engines
    against a small in-memory fake catalog - no real control modules or live
    tenant connection needed.
#>

BeforeAll {
    Import-Module (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1') -Force
}

Describe 'Compliance diffing (Test-BaselineCompliance / Compare-BaselineValueDeep)' {

    It 'reports compliant for identical scalars' {
        Test-BaselineCompliance -CurrentValue 'Direct' -DesiredValue 'Direct' | Should -Be $true
    }

    It 'reports non-compliant for different scalars' {
        Test-BaselineCompliance -CurrentValue 'AnonymousAccess' -DesiredValue 'Direct' | Should -Be $false
    }

    It 'reports Unknown (null) when the current value could not be read' {
        Test-BaselineCompliance -CurrentValue $null -DesiredValue 'Direct' | Should -Be $null
    }

    It 'deep-compares nested objects regardless of property order' {
        $current = [pscustomobject]@{ b = 2; a = 1 }
        $desired = [pscustomobject]@{ a = 1; b = 2 }
        Test-BaselineCompliance -CurrentValue $current -DesiredValue $desired | Should -Be $true
    }

    It 'detects a difference nested inside an object' {
        $current = [pscustomobject]@{ a = 1; b = 3 }
        $desired = [pscustomobject]@{ a = 1; b = 2 }
        Test-BaselineCompliance -CurrentValue $current -DesiredValue $desired | Should -Be $false
    }

    It 'compares arrays element-by-element' {
        Compare-BaselineValueDeep -Left @('a', 'b') -Right @('a', 'b') | Should -Be $true
        Compare-BaselineValueDeep -Left @('a', 'b') -Right @('a', 'c') | Should -Be $false
        Compare-BaselineValueDeep -Left @('a') -Right @('a', 'b') | Should -Be $false
    }

    It 'evaluates Range compliance mode against min/max' {
        Test-BaselineCompliance -CurrentValue 3 -DesiredValue @{ min = 2; max = 4 } -ComplianceMode Range | Should -Be $true
        Test-BaselineCompliance -CurrentValue 6 -DesiredValue @{ min = 2; max = 4 } -ComplianceMode Range | Should -Be $false
        Test-BaselineCompliance -CurrentValue 1 -DesiredValue @{ min = 2; max = 4 } -ComplianceMode Range | Should -Be $false
    }

    Context 'Hashtable values (not just PSCustomObject) - regression test for a real infinite-recursion bug' {
        # A Hashtable is ALSO [System.Collections.IEnumerable] in .NET, same as an array.
        # Comparing two hashtables used to recurse forever: @($hashtable) just wraps the
        # SAME hashtable back into a 1-element array containing itself, so the recursive
        # per-element compare called Compare-BaselineValueDeep with the identical two
        # hashtable arguments again, forever, until a ScriptCallDepthException. Caught by
        # a real Pester run on a live workstation (EntraID-AuthMethodsHardening's
        # systemCredentialPreferences test, using a nested hashtable, hung for 5+ minutes
        # before erroring). Timeout-guarded here so a regression fails fast, not by hanging.

        It 'compares two equal hashtables as compliant, quickly (does not recurse forever)' {
            $job = Start-Job -ScriptBlock {
                param($modulePath)
                Import-Module $modulePath -Force
                Compare-BaselineValueDeep -Left @{ state = 'enabled' } -Right @{ state = 'enabled' }
            } -ArgumentList (Join-Path $PSScriptRoot '../modules/BaselineCore.psm1')
            $completed = Wait-Job -Job $job -Timeout 10
            if (-not $completed) { Stop-Job -Job $job; Remove-Job -Job $job -Force; throw "Compare-BaselineValueDeep did not return within 10 seconds - infinite recursion regression." }
            $result = Receive-Job -Job $job
            Remove-Job -Job $job -Force
            $result | Should -Be $true
        }

        It 'compares two different hashtables as non-compliant' {
            Compare-BaselineValueDeep -Left @{ state = 'enabled' } -Right @{ state = 'disabled' } | Should -Be $false
        }

        It 'compares a hashtable and a PSCustomObject with equivalent content as equal' {
            Compare-BaselineValueDeep -Left @{ state = 'enabled' } -Right ([pscustomobject]@{ state = 'enabled' }) | Should -Be $true
        }

        It 'compares a hashtable nested inside a PSCustomObject (the exact shape that triggered the bug)' {
            $same = [pscustomobject]@{ authenticatorEnabled = $true; systemCredentialPreferences = @{ state = 'enabled' } }
            Compare-BaselineValueDeep -Left $same -Right $same | Should -Be $true
        }

        It 'still compares an array of hashtables correctly (e.g. MfaRegistrationCampaign includeTargets)' {
            $targets1 = @(@{ targetType = 'group'; id = 'all_users' })
            $targets2 = @(@{ targetType = 'group'; id = 'all_users' })
            $targets3 = @(@{ targetType = 'group'; id = 'different' })
            Compare-BaselineValueDeep -Left $targets1 -Right $targets2 | Should -Be $true
            Compare-BaselineValueDeep -Left $targets1 -Right $targets3 | Should -Be $false
        }
    }
}

Describe 'Invoke-BaselineControlAudit classification' {

    BeforeEach {
        function global:Get-Fake-OkState { [pscustomobject]@{ Id = 'Fake-Ok'; Value = 'Direct' } }
        function global:Set-Fake-OkState { param($DesiredValue, $CurrentValue) [pscustomobject]@{ Id = 'Fake-Ok'; Status = 'Success'; PreviousValue = $CurrentValue; AppliedValue = $DesiredValue; Message = '' } }

        # Tracks state via a script-scoped variable, updated by Set-, rather than a
        # hardcoded value that never changes - Invoke-BaselineControlApply now calls
        # Test-BaselineApplyOutcome after any successful Set-, which re-reads via
        # Get-Fake-DriftState to confirm the change landed. A static Get- that always
        # returns the old value (as this fixture did before) makes that read-back
        # never match, misclassifying a genuinely successful Set- as
        # Applied-PendingConfirmation - not what this test is checking.
        $script:FakeDriftValue = 'AnonymousAccess'
        function global:Get-Fake-DriftState { [pscustomobject]@{ Id = 'Fake-Drift'; Value = $script:FakeDriftValue } }
        function global:Set-Fake-DriftState { param($DesiredValue, $CurrentValue) $script:FakeDriftValue = $DesiredValue; [pscustomobject]@{ Id = 'Fake-Drift'; Status = 'Success'; PreviousValue = $CurrentValue; AppliedValue = $DesiredValue; Message = '' } }

        function global:Get-Fake-ManualState { [pscustomobject]@{ Id = 'Fake-Manual'; Value = $null } }
        function global:Set-Fake-ManualState { param($DesiredValue, $CurrentValue) [pscustomobject]@{ Id = 'Fake-Manual'; Status = 'Skipped-Manual'; PreviousValue = $CurrentValue; AppliedValue = $null; Message = 'See GUI.' } }

        function global:Get-Fake-ErrorState { throw 'Simulated read failure (e.g. insufficient permissions).' }
        function global:Set-Fake-ErrorState { param($DesiredValue, $CurrentValue) throw 'Should never be called.' }

        $script:Catalog = @(
            [pscustomobject]@{ Id = 'Fake-Ok'; Workload = 'SharePointOnline'; Connection = 'SharePointOnline'; Automatable = $true; DesiredValue = 'Direct'; Description = 'ok'; ComplianceMode = 'Equality'; ManualInstructions = ''; RequiresPopulatedFields = @(); GetCommand = 'Get-Fake-OkState'; SetCommand = 'Set-Fake-OkState' }
            [pscustomobject]@{ Id = 'Fake-Drift'; Workload = 'SharePointOnline'; Connection = 'SharePointOnline'; Automatable = $true; DesiredValue = 'Direct'; Description = 'drift'; ComplianceMode = 'Equality'; ManualInstructions = ''; RequiresPopulatedFields = @(); GetCommand = 'Get-Fake-DriftState'; SetCommand = 'Set-Fake-DriftState' }
            [pscustomobject]@{ Id = 'Fake-Manual'; Workload = 'EntraID'; Connection = 'Graph'; Automatable = $false; DesiredValue = $true; Description = 'manual'; ComplianceMode = 'Equality'; ManualInstructions = 'Do it by hand.'; RequiresPopulatedFields = @(); GetCommand = 'Get-Fake-ManualState'; SetCommand = 'Set-Fake-ManualState' }
            [pscustomobject]@{ Id = 'Fake-Error'; Workload = 'EntraID'; Connection = 'Graph'; Automatable = $true; DesiredValue = $true; Description = 'errors on read'; ComplianceMode = 'Equality'; ManualInstructions = ''; RequiresPopulatedFields = @(); GetCommand = 'Get-Fake-ErrorState'; SetCommand = 'Set-Fake-ErrorState' }
        )
    }

    It 'classifies a compliant control correctly' {
        $audit = Invoke-BaselineControlAudit -Catalog $script:Catalog
        ($audit | Where-Object Id -eq 'Fake-Ok').Compliant | Should -Be $true
    }

    It 'classifies a drifted control as non-compliant' {
        $audit = Invoke-BaselineControlAudit -Catalog $script:Catalog
        ($audit | Where-Object Id -eq 'Fake-Drift').Compliant | Should -Be $false
    }

    It 'classifies a control with no readable value as Unknown (null), not an error' {
        $audit = Invoke-BaselineControlAudit -Catalog $script:Catalog
        $manual = $audit | Where-Object Id -eq 'Fake-Manual'
        $manual.Compliant | Should -Be $null
        $manual.Error | Should -BeNullOrEmpty
    }

    It 'captures a read failure as an Error, distinct from non-compliance' {
        $audit = Invoke-BaselineControlAudit -Catalog $script:Catalog
        $errored = $audit | Where-Object Id -eq 'Fake-Error'
        $errored.Compliant | Should -Be $null
        $errored.Error | Should -Match 'Simulated read failure'
    }

    It 'skips a Skipped-Manual control during Apply without calling its Set- function, using config manualInstructions' {
        $audit = Invoke-BaselineControlAudit -Catalog $script:Catalog
        $logPath = [System.IO.Path]::GetTempFileName()
        try {
            $applyResults = Invoke-BaselineControlApply -Catalog $script:Catalog -AuditResults $audit -ChangeLogPath $logPath
            $manualResult = $applyResults | Where-Object Id -eq 'Fake-Manual'
            $manualResult.Status | Should -Be 'Skipped-Manual'
            $manualResult.Message | Should -Be 'Do it by hand.'

            $okResult = $applyResults | Where-Object Id -eq 'Fake-Ok'
            $okResult.Status | Should -Be 'Skipped-AlreadyCompliant'

            $driftResult = $applyResults | Where-Object Id -eq 'Fake-Drift'
            $driftResult.Status | Should -Be 'Success'
            $driftResult.AppliedValue | Should -Be 'Direct'

            $errorResult = $applyResults | Where-Object Id -eq 'Fake-Error'
            $errorResult.Status | Should -Be 'Failed'
        }
        finally {
            Remove-Item $logPath -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Restore-mode logic' {

    BeforeEach {
        $script:CapturedDesiredValue = $null
        function global:Get-Fake-RestoreState { [pscustomobject]@{ Id = 'Fake-Restore'; Value = 'ExternalUserAndGuestSharing' } }
        function global:Set-Fake-RestoreState {
            param($DesiredValue, $CurrentValue)
            $script:CapturedDesiredValue = $DesiredValue
            [pscustomobject]@{ Id = 'Fake-Restore'; Status = 'Success'; PreviousValue = $CurrentValue; AppliedValue = $DesiredValue; Message = 'Restored.' }
        }

        # The *current live config's* desired value is deliberately different from what
        # the snapshot recorded, to prove Restore uses the snapshot's value, not this one.
        $script:Catalog = @(
            [pscustomobject]@{ Id = 'Fake-Restore'; Workload = 'SharePointOnline'; Connection = 'SharePointOnline'; Automatable = $true; DesiredValue = 'ExternalUserSharingOnly'; Description = 'restore test'; ComplianceMode = 'Equality'; ManualInstructions = ''; RequiresPopulatedFields = @(); GetCommand = 'Get-Fake-RestoreState'; SetCommand = 'Set-Fake-RestoreState' }
        )

        $script:SnapshotControls = @(
            [pscustomobject]@{ id = 'Fake-Restore'; workload = 'SharePointOnline'; currentValue = 'ExternalUserAndGuestSharing'; desiredValue = 'ExternalUserSharingOnly' }
        )
    }

    It 'calls Set- with the snapshot''s recorded currentValue, not the live config''s desiredValue' {
        $logPath = [System.IO.Path]::GetTempFileName()
        try {
            $results = Invoke-BaselineControlRestore -Catalog $script:Catalog -SnapshotControls $script:SnapshotControls -ChangeLogPath $logPath
            $script:CapturedDesiredValue | Should -Be 'ExternalUserAndGuestSharing'
            $script:CapturedDesiredValue | Should -Not -Be 'ExternalUserSharingOnly'
            ($results | Where-Object Id -eq 'Fake-Restore').Status | Should -Be 'Success'
        }
        finally {
            Remove-Item $logPath -ErrorAction SilentlyContinue
        }
    }

    It 'marks a control absent from the current catalog as Failed rather than silently skipping it' {
        $orphanSnapshot = @([pscustomobject]@{ id = 'Fake-DoesNotExistAnymore'; workload = 'SharePointOnline'; currentValue = 'x'; desiredValue = 'y' })
        $logPath = [System.IO.Path]::GetTempFileName()
        try {
            $results = Invoke-BaselineControlRestore -Catalog $script:Catalog -SnapshotControls $orphanSnapshot -ChangeLogPath $logPath
            $results[0].Status | Should -Be 'Failed'
            $results[0].Message | Should -Match 'no control'
        }
        finally {
            Remove-Item $logPath -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Test-BaselineApplyOutcome (generic read-back-and-classify helper)' {

    It 'reports Success/Confirmed when the first read-back already matches' {
        $script:callCount = 0
        function global:Get-Fake-Outcome1State { $script:callCount++; [pscustomobject]@{ Value = $true } }
        $r = Test-BaselineApplyOutcome -GetCommand 'Get-Fake-Outcome1State' -DesiredValue $true -RetryDelayMilliseconds 0
        $r.Status | Should -Be 'Success'
        $r.Confirmed | Should -Be $true
        $script:callCount | Should -Be 1
    }

    It 'reports Success/Confirmed when the first read-back mismatches but the retry matches' {
        $script:callCount2 = 0
        function global:Get-Fake-Outcome2State {
            $script:callCount2++
            if ($script:callCount2 -eq 1) { [pscustomobject]@{ Value = $false } } else { [pscustomobject]@{ Value = $true } }
        }
        $r = Test-BaselineApplyOutcome -GetCommand 'Get-Fake-Outcome2State' -DesiredValue $true -RetryDelayMilliseconds 0
        $r.Status | Should -Be 'Success'
        $r.Confirmed | Should -Be $true
        $script:callCount2 | Should -Be 2
    }

    It 'reports Applied-PendingConfirmation when both reads mismatch and the control is not flagged mechanismPossiblyDeprecated' {
        function global:Get-Fake-Outcome3State { [pscustomobject]@{ Value = $false } }
        $r = Test-BaselineApplyOutcome -GetCommand 'Get-Fake-Outcome3State' -DesiredValue $true -RetryDelayMilliseconds 0
        $r.Status | Should -Be 'Applied-PendingConfirmation'
        $r.Confirmed | Should -Be $false
    }

    It 'reports MechanismPossiblyDeprecated when both reads mismatch and -MechanismPossiblyDeprecated is set' {
        function global:Get-Fake-Outcome4State { [pscustomobject]@{ Value = $false } }
        $r = Test-BaselineApplyOutcome -GetCommand 'Get-Fake-Outcome4State' -DesiredValue $true -MechanismPossiblyDeprecated -RetryDelayMilliseconds 0
        $r.Status | Should -Be 'MechanismPossiblyDeprecated'
        $r.Confirmed | Should -Be $false
    }

    It 'a genuine thrown error from Set- stays Failed and never reaches the read-back helper at all' {
        function global:Get-Fake-Outcome5State { [pscustomobject]@{ Value = $false } }
        function global:Set-Fake-Outcome5State { param($DesiredValue, $CurrentValue) throw 'Simulated Graph API failure (e.g. insufficient permissions).' }

        $catalog = @(
            [pscustomobject]@{ Id = 'Fake-Outcome5'; Workload = 'EntraID'; Connection = 'Graph'; Automatable = $true; DesiredValue = $true; Description = 'x'; ComplianceMode = 'Equality'; ManualInstructions = ''; RequiresPopulatedFields = @(); GetCommand = 'Get-Fake-Outcome5State'; SetCommand = 'Set-Fake-Outcome5State' }
        )
        $audit = Invoke-BaselineControlAudit -Catalog $catalog
        $logPath = [System.IO.Path]::GetTempFileName()
        try {
            $results = Invoke-BaselineControlApply -Catalog $catalog -AuditResults $audit -ChangeLogPath $logPath -ApplyOutcomeRetryDelayMilliseconds 0
            $results[0].Status | Should -Be 'Failed'
            $results[0].Message | Should -Match 'Simulated Graph API failure'
        }
        finally {
            Remove-Item $logPath -ErrorAction SilentlyContinue
        }
    }

    It 'Invoke-BaselineControlApply upgrades a Success result to MechanismPossiblyDeprecated end-to-end when the catalog flag is set and the read-back never confirms' {
        function global:Get-Fake-Outcome6State { [pscustomobject]@{ Value = $false } }
        function global:Set-Fake-Outcome6State { param($DesiredValue, $CurrentValue) [pscustomobject]@{ Id = 'Fake-Outcome6'; Status = 'Success'; PreviousValue = $CurrentValue; AppliedValue = $DesiredValue; Message = 'API reported success.' } }

        $catalog = @(
            [pscustomobject]@{ Id = 'Fake-Outcome6'; Workload = 'SharePointOnline'; Connection = 'SharePointOnline'; Automatable = $true; DesiredValue = $true; Description = 'x'; ComplianceMode = 'Equality'; ManualInstructions = ''; RequiresPopulatedFields = @(); MechanismPossiblyDeprecated = $true; GetCommand = 'Get-Fake-Outcome6State'; SetCommand = 'Set-Fake-Outcome6State' }
        )
        $audit = Invoke-BaselineControlAudit -Catalog $catalog
        $logPath = [System.IO.Path]::GetTempFileName()
        try {
            $results = Invoke-BaselineControlApply -Catalog $catalog -AuditResults $audit -ChangeLogPath $logPath -ApplyOutcomeRetryDelayMilliseconds 0
            $results[0].Status | Should -Be 'MechanismPossiblyDeprecated'

            $logEntries = Get-Content $logPath | ForEach-Object { $_ | ConvertFrom-Json }
            ($logEntries | Where-Object id -eq 'Fake-Outcome6').result | Should -Be 'MechanismPossiblyDeprecated'
        }
        finally {
            Remove-Item $logPath -ErrorAction SilentlyContinue
        }
    }

    It 'does not throw StrictMode errors against a hand-built catalog entry with no MechanismPossiblyDeprecated property at all' {
        function global:Get-Fake-Outcome7State { [pscustomobject]@{ Value = 'Direct' } }
        function global:Set-Fake-Outcome7State { param($DesiredValue, $CurrentValue) [pscustomobject]@{ Id = 'Fake-Outcome7'; Status = 'Success'; PreviousValue = $CurrentValue; AppliedValue = $DesiredValue; Message = '' } }

        # Deliberately omits MechanismPossiblyDeprecated (and Tier/ForceCreateDespiteOverlap),
        # mirroring every other hand-built fixture catalog in this file - this module runs
        # under Set-StrictMode -Version Latest, so a naive unguarded property access would throw.
        $catalog = @(
            [pscustomobject]@{ Id = 'Fake-Outcome7'; Workload = 'SharePointOnline'; Connection = 'SharePointOnline'; Automatable = $true; DesiredValue = 'Direct'; Description = 'x'; ComplianceMode = 'Equality'; ManualInstructions = ''; RequiresPopulatedFields = @(); GetCommand = 'Get-Fake-Outcome7State'; SetCommand = 'Set-Fake-Outcome7State' }
        )
        $audit = Invoke-BaselineControlAudit -Catalog $catalog
        $logPath = [System.IO.Path]::GetTempFileName()
        try {
            { Invoke-BaselineControlApply -Catalog $catalog -AuditResults $audit -ChangeLogPath $logPath -ApplyOutcomeRetryDelayMilliseconds 0 } | Should -Not -Throw
        }
        finally {
            Remove-Item $logPath -ErrorAction SilentlyContinue
        }
    }
}

Describe 'Manual-review highlighting in reports' {

    BeforeAll {
        $script:ManualReviewAuditResults = @(
            [pscustomobject]@{ Id = 'EntraID-AutoControl'; Workload = 'EntraID'; Description = 'auto'; Automatable = $true; CurrentValue = $true; DesiredValue = $true; Compliant = $true; ManualInstructions = ''; Detail = $null; Error = $null }
            [pscustomobject]@{ Id = 'EntraID-GaNotLocalAdminOnJoin'; Workload = 'EntraID'; Description = 'manual1'; Automatable = $false; CurrentValue = $null; DesiredValue = $true; Compliant = $null; ManualInstructions = 'Entra admin center > Identity > Devices > Device settings'; Detail = 'no api'; Error = $null }
            [pscustomobject]@{ Id = 'M365AdminCenter-SwayExternalSharing'; Workload = 'M365AdminCenter'; Description = 'manual2'; Automatable = $false; CurrentValue = $null; DesiredValue = $false; Compliant = $null; ManualInstructions = 'M365 admin center > Org settings > Sway'; Detail = 'no api'; Error = $null }
        )
    }

    Context 'Export-BaselineMarkdownReport' {
        It 'adds a top-of-report summary section listing every manual-review control with its instructions' {
            $path = [System.IO.Path]::GetTempFileName()
            try {
                Export-BaselineMarkdownReport -AuditResults $script:ManualReviewAuditResults -Path $path -Title 'Test' | Out-Null
                $content = Get-Content $path -Raw
                $content | Should -Match 'Manual review required: 2'
                $content | Should -Match 'Manual review required\r?\n'
                $content | Should -Match '\*\*EntraID-GaNotLocalAdminOnJoin\*\*.*Device settings'
                $content | Should -Match '\*\*M365AdminCenter-SwayExternalSharing\*\*.*Org settings'
            }
            finally {
                Remove-Item $path -ErrorAction SilentlyContinue
            }
        }

        It 'marks each manual-review control''s row (not just automatable ones) with a warning prefix on its Id' {
            $path = [System.IO.Path]::GetTempFileName()
            try {
                Export-BaselineMarkdownReport -AuditResults $script:ManualReviewAuditResults -Path $path -Title 'Test' | Out-Null
                $content = Get-Content $path -Raw
                $content | Should -Match '\*\*EntraID-GaNotLocalAdminOnJoin\*\* \|'
                $content | Should -Not -Match '\*\*EntraID-AutoControl\*\*'
            }
            finally {
                Remove-Item $path -ErrorAction SilentlyContinue
            }
        }

        It 'omits the summary section entirely when there are no manual-review controls' {
            $path = [System.IO.Path]::GetTempFileName()
            try {
                $onlyAuto = @($script:ManualReviewAuditResults | Where-Object Automatable)
                Export-BaselineMarkdownReport -AuditResults $onlyAuto -Path $path -Title 'Test' | Out-Null
                $content = Get-Content $path -Raw
                $content | Should -Not -Match 'Manual review required\r?\n'
                $content | Should -Match 'Manual review required: 0'
            }
            finally {
                Remove-Item $path -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'Export-BaselineHtmlReport' {
        It 'adds a manual-summary callout div listing every manual-review control' {
            $path = [System.IO.Path]::GetTempFileName()
            try {
                Export-BaselineHtmlReport -AuditResults $script:ManualReviewAuditResults -Path $path -Title 'Test' | Out-Null
                $content = Get-Content $path -Raw
                $content | Should -Match '<div class="manual-summary">'
                $content | Should -Match 'EntraID-GaNotLocalAdminOnJoin'
                $content | Should -Match 'M365AdminCenter-SwayExternalSharing'
            }
            finally {
                Remove-Item $path -ErrorAction SilentlyContinue
            }
        }

        It 'applies the manual-cell CSS class to exactly the non-automatable rows'' Result/Notes cell, not the whole row' {
            $path = [System.IO.Path]::GetTempFileName()
            try {
                Export-BaselineHtmlReport -AuditResults $script:ManualReviewAuditResults -Path $path -Title 'Test' | Out-Null
                $content = Get-Content $path -Raw
                $manualCellCount = ([regex]::Matches($content, 'class="manual-cell"')).Count
                $manualCellCount | Should -Be 2
                # Cell-level, not row-level: no bare "<tr class=" left over from the old
                # whole-row highlight.
                $content | Should -Not -Match '<tr class='
            }
            finally {
                Remove-Item $path -ErrorAction SilentlyContinue
            }
        }

        It 'applies the compliant-yes CSS class to exactly the compliant controls'' Compliant cell' {
            $path = [System.IO.Path]::GetTempFileName()
            try {
                Export-BaselineHtmlReport -AuditResults $script:ManualReviewAuditResults -Path $path -Title 'Test' | Out-Null
                $content = Get-Content $path -Raw
                $compliantYesCount = ([regex]::Matches($content, 'class="compliant-yes"')).Count
                $compliantYesCount | Should -Be 1
            }
            finally {
                Remove-Item $path -ErrorAction SilentlyContinue
            }
        }

        It 'omits the manual-summary div and manual-cell class entirely when there are no manual-review controls' {
            $path = [System.IO.Path]::GetTempFileName()
            try {
                $onlyAuto = @($script:ManualReviewAuditResults | Where-Object Automatable)
                Export-BaselineHtmlReport -AuditResults $onlyAuto -Path $path -Title 'Test' | Out-Null
                $content = Get-Content $path -Raw
                $content | Should -Not -Match '<div class="manual-summary">'
                $content | Should -Not -Match 'class="manual-cell"'
            }
            finally {
                Remove-Item $path -ErrorAction SilentlyContinue
            }
        }

        It 'uses a fixed table layout with percentage column widths so the table fits the viewport instead of auto-growing wide columns' {
            $path = [System.IO.Path]::GetTempFileName()
            try {
                Export-BaselineHtmlReport -AuditResults $script:ManualReviewAuditResults -Path $path -Title 'Test' | Out-Null
                $content = Get-Content $path -Raw
                $content | Should -Match 'table-layout:\s*fixed'
                $content | Should -Match '<colgroup>'
                $content | Should -Match 'overflow-wrap:\s*anywhere'
                $content | Should -Match 'class="table-wrap"'
                $content | Should -Match '@media \(max-width: 900px\)'
            }
            finally {
                Remove-Item $path -ErrorAction SilentlyContinue
            }
        }

        It 'uses a 9-column colgroup for an Apply/Restore report and an 8-column colgroup for a plain Audit report' {
            $auditOnlyPath = [System.IO.Path]::GetTempFileName()
            $applyPath = [System.IO.Path]::GetTempFileName()
            try {
                Export-BaselineHtmlReport -AuditResults $script:ManualReviewAuditResults -Path $auditOnlyPath -Title 'Test' | Out-Null
                $applyResults = @($script:ManualReviewAuditResults | ForEach-Object { [pscustomobject]@{ Id = $_.Id; Status = 'Success'; Message = '' } })
                Export-BaselineHtmlReport -AuditResults $script:ManualReviewAuditResults -Path $applyPath -Title 'Test' -ApplyResults $applyResults | Out-Null

                $auditOnlyColCount = ([regex]::Matches((Get-Content $auditOnlyPath -Raw), '<col ')).Count
                $applyColCount = ([regex]::Matches((Get-Content $applyPath -Raw), '<col ')).Count
                $auditOnlyColCount | Should -Be 8
                $applyColCount | Should -Be 9
            }
            finally {
                Remove-Item $auditOnlyPath, $applyPath -ErrorAction SilentlyContinue
            }
        }
    }
}

Describe 'Snapshot schema version enforcement' {

    It 'refuses to load a snapshot with an unsupported snapshotSchemaVersion' {
        $badSnapshot = [pscustomobject]@{
            snapshotSchemaVersion = '99.0'
            capturedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
            mode = 'Audit'
            baselineConfigSchemaVersion = '1.0'
            controls = @([pscustomobject]@{ id = 'x'; workload = 'EntraID'; currentValue = $true; desiredValue = $true })
        }
        $path = [System.IO.Path]::GetTempFileName()
        try {
            $badSnapshot | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $path
            { Import-BaselineSnapshot -Path $path } | Should -Throw '*snapshotSchemaVersion*'
        }
        finally {
            Remove-Item $path -ErrorAction SilentlyContinue
        }
    }
}
