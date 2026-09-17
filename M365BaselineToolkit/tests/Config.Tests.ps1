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
