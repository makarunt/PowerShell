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
}

Describe 'Invoke-BaselineControlAudit classification' {

    BeforeEach {
        function global:Get-Fake-OkState { [pscustomobject]@{ Id = 'Fake-Ok'; Value = 'Direct' } }
        function global:Set-Fake-OkState { param($DesiredValue, $CurrentValue) [pscustomobject]@{ Id = 'Fake-Ok'; Status = 'Success'; PreviousValue = $CurrentValue; AppliedValue = $DesiredValue; Message = '' } }

        function global:Get-Fake-DriftState { [pscustomobject]@{ Id = 'Fake-Drift'; Value = 'AnonymousAccess' } }
        function global:Set-Fake-DriftState { param($DesiredValue, $CurrentValue) [pscustomobject]@{ Id = 'Fake-Drift'; Status = 'Success'; PreviousValue = $CurrentValue; AppliedValue = $DesiredValue; Message = '' } }

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
