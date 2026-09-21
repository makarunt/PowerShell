<#
    Config.Tests.ps1

    Validates the baseline config schema/semantic validation logic. Runs entirely
    against local files and in-memory objects - no live tenant connection needed.
#>

BeforeAll {
    $script:ModulesDir = Join-Path $PSScriptRoot '../modules'
    $script:ConfigDir = Join-Path $PSScriptRoot '../config'
    Import-Module (Join-Path $script:ModulesDir 'BaselineCore.psm1') -Force

    $script:ValidConfigPath = Join-Path $script:ConfigDir 'baseline.config.json'
    $script:SchemaPath = Join-Path $script:ConfigDir 'baseline.config.schema.json'

    function New-TempConfigFile {
        param([Parameter(Mandatory)][object]$ConfigObject)
        $path = [System.IO.Path]::GetTempFileName()
        $ConfigObject | ConvertTo-Json -Depth 25 | Set-Content -LiteralPath $path -Encoding utf8
        return $path
    }
}

Describe 'Baseline config validation' {

    Context 'The shipped seed config' {
        It 'passes validation as-is' {
            { Import-BaselineConfig -Path $script:ValidConfigPath -SchemaPath $script:SchemaPath } | Should -Not -Throw
        }

        It 'defines at least one control for every workload' {
            $config = Import-BaselineConfig -Path $script:ValidConfigPath -SchemaPath $script:SchemaPath
            $workloads = $config.controls.workload | Select-Object -Unique
            $workloads | Should -Contain 'EntraID'
            $workloads | Should -Contain 'ExchangeOnline'
            $workloads | Should -Contain 'Teams'
            $workloads | Should -Contain 'SharePointOnline'
            $workloads | Should -Contain 'M365AdminCenter'
        }

        It 'flags SharePointOnline-AzureADB2BIntegration as mechanismPossiblyDeprecated' {
            $config = Import-BaselineConfig -Path $script:ValidConfigPath -SchemaPath $script:SchemaPath
            $control = $config.controls | Where-Object { $_.id -eq 'SharePointOnline-AzureADB2BIntegration' }
            $control | Should -Not -BeNullOrEmpty
            $control.mechanismPossiblyDeprecated | Should -Be $true
        }

        It 'requires reviewers to be populated before EntraID-AdminConsentWorkflow can be applied' {
            $config = Import-BaselineConfig -Path $script:ValidConfigPath -SchemaPath $script:SchemaPath
            $catalog = Get-BaselineControlCatalog -Config $config
            $readinessErrors = Test-BaselineApplyReadiness -Catalog $catalog
            ($readinessErrors -join "`n") | Should -Match 'EntraID-AdminConsentWorkflow.*reviewers is empty'
        }
    }

    Context 'M365AdminCenter workload' {
        It 'accepts M365AdminCenter as a valid workload enum value' {
            $config = [pscustomobject]@{
                schemaVersion = '1.0'
                baselineSource = 'test'
                controls = @(
                    [pscustomobject]@{
                        id = 'M365AdminCenter-Test'
                        workload = 'M365AdminCenter'
                        enabled = $true
                        automatable = $false
                        desiredValue = $true
                        description = 'A test control on the M365AdminCenter workload.'
                        manualInstructions = 'Change it by hand.'
                    }
                )
            }
            $goodPath = New-TempConfigFile -ConfigObject $config
            try {
                { Import-BaselineConfig -Path $goodPath -SchemaPath $script:SchemaPath } | Should -Not -Throw
            }
            finally {
                Remove-Item $goodPath -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'A config with a bad enum value' {
        It 'fails validation' {
            $config = Get-Content $script:ValidConfigPath -Raw | ConvertFrom-Json -Depth 25
            $config.controls[0].workload = 'NotARealWorkload'
            $badPath = New-TempConfigFile -ConfigObject $config
            try {
                { Import-BaselineConfig -Path $badPath -SchemaPath $script:SchemaPath } | Should -Throw
            }
            finally {
                Remove-Item $badPath -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'A config with a missing required field' {
        It 'fails validation when a control is missing desiredValue' {
            $config = Get-Content $script:ValidConfigPath -Raw | ConvertFrom-Json -Depth 25
            $control = $config.controls[0] | Select-Object * -ExcludeProperty desiredValue
            $config.controls[0] = $control
            $badPath = New-TempConfigFile -ConfigObject $config
            try {
                { Import-BaselineConfig -Path $badPath -SchemaPath $script:SchemaPath } | Should -Throw
            }
            finally {
                Remove-Item $badPath -ErrorAction SilentlyContinue
            }
        }

        It 'fails validation when an automatable:false control has no manualInstructions' {
            $config = [pscustomobject]@{
                schemaVersion = '1.0'
                baselineSource = 'test'
                controls = @(
                    [pscustomobject]@{
                        id = 'Test-Manual'
                        workload = 'EntraID'
                        enabled = $true
                        automatable = $false
                        desiredValue = $true
                        description = 'A manual-only control with no instructions.'
                    }
                )
            }
            $badPath = New-TempConfigFile -ConfigObject $config
            try {
                { Import-BaselineConfig -Path $badPath -SchemaPath $script:SchemaPath } | Should -Throw '*manualInstructions*'
            }
            finally {
                Remove-Item $badPath -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'A config referencing an unknown control id' {
        It 'fails catalog validation when a control has no matching Get-/Set- functions' {
            $config = [pscustomobject]@{
                schemaVersion = '1.0'
                baselineSource = 'test'
                controls = @(
                    [pscustomobject]@{
                        id = 'TotallyMade-UpControl'
                        workload = 'EntraID'
                        enabled = $true
                        automatable = $true
                        desiredValue = $true
                        description = 'Does not exist in any control module.'
                    }
                )
            }
            { Get-BaselineControlCatalog -Config $config -AvailableFunctions @() } | Should -Throw '*TotallyMade-UpControl*'
        }
    }

    Context 'A config with duplicate control ids' {
        It 'fails semantic validation' {
            $config = Get-Content $script:ValidConfigPath -Raw | ConvertFrom-Json -Depth 25
            $dupe = $config.controls[0] | Select-Object *
            $config.controls += $dupe
            $badPath = New-TempConfigFile -ConfigObject $config
            try {
                { Import-BaselineConfig -Path $badPath -SchemaPath $script:SchemaPath } | Should -Throw '*duplicate*'
            }
            finally {
                Remove-Item $badPath -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'A config with an unsupported schemaVersion' {
        It 'fails validation' {
            $config = Get-Content $script:ValidConfigPath -Raw | ConvertFrom-Json -Depth 25
            $config.schemaVersion = '99.0'
            $badPath = New-TempConfigFile -ConfigObject $config
            try {
                { Import-BaselineConfig -Path $badPath -SchemaPath $script:SchemaPath } | Should -Throw '*schemaVersion*'
            }
            finally {
                Remove-Item $badPath -ErrorAction SilentlyContinue
            }
        }
    }

    Context 'Range compliance mode validation' {
        It 'fails when complianceMode is Range but desiredValue has no min/max' {
            $config = [pscustomobject]@{
                schemaVersion = '1.0'
                baselineSource = 'test'
                controls = @(
                    [pscustomobject]@{
                        id = 'Test-Range'
                        workload = 'EntraID'
                        enabled = $true
                        automatable = $true
                        desiredValue = 5
                        description = 'Bad range control.'
                        complianceMode = 'Range'
                    }
                )
            }
            $badPath = New-TempConfigFile -ConfigObject $config
            try {
                { Import-BaselineConfig -Path $badPath -SchemaPath $script:SchemaPath } | Should -Throw '*Range*'
            }
            finally {
                Remove-Item $badPath -ErrorAction SilentlyContinue
            }
        }
    }
}
