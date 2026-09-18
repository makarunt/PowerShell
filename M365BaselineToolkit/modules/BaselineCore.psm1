#Requires -Version 7.0
<#
    BaselineCore.psm1

    Orchestration engine for the M365 Baseline Toolkit: config loading/validation,
    connection management, the control catalog, compliance diffing, backup/restore,
    and reporting. This module contains no tenant-specific business logic - that
    lives in the per-workload control modules (EntraIdControls.psm1, etc). This
    module never reads executable strings out of the config file; the config file
    is data only.
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# Schema version this build of the toolkit understands for the config file.
$script:SupportedConfigSchemaVersion = '1.0'

# Schema version this build of the toolkit writes/reads for snapshot (backup) files.
$script:SnapshotSchemaVersion = '1.0'

# Maps a control's connection requirement to the PowerShell module that provides it.
# Key order here is also the canonical CONNECTION order (see Get-BaselineConnectionOrder
# below): Microsoft.Graph must connect before ExchangeOnlineManagement in the same
# PowerShell process. Both modules bundle their own copy of MSAL (Microsoft.Identity.Client)
# and, once one module's copy is loaded into the process, .NET keeps using that exact
# version for the rest of the session - if Exchange Online's older bundled MSAL loads
# first, Microsoft.Graph's later Connect-MgGraph call fails with a MissingMethodException
# ("Method not found: ...WithLogging...") even though the account has every permission it
# needs. Connecting to Graph first sidesteps this well-documented cross-module conflict.
$script:WorkloadModuleMap = [ordered]@{
    Graph            = 'Microsoft.Graph'
    ExchangeOnline   = 'ExchangeOnlineManagement'
    Teams            = 'MicrosoftTeams'
    SharePointOnline = 'Microsoft.Online.SharePoint.PowerShell'
}

# Narrowest Graph delegated scopes that cover the EntraID controls in this toolkit,
# plus what's needed for license-gated controls (Organization.Read.All, for
# Test-TenantServicePlan/Get-MgSubscribedSku) and the Conditional Access module
# (Policy.ReadWrite.ConditionalAccess, Group.ReadWrite.All for the emergency-access
# group, Application.Read.All to resolve the Azure Management service principal).
# Requested unconditionally rather than computed per-run from which controls are
# enabled: granting a scope costs nothing on a tenant that can't use the feature
# behind it (e.g. Policy.ReadWrite.ConditionalAccess consents fine on Entra ID Free,
# it's ACTUALLY reading/writing a CA policy that the license gate in
# ConditionalAccessControls.psm1 prevents on Free) and keeps the connection logic
# simple - only the license gate decides what actually happens, never the scope list.
#
# Policy.Read.All is required separately from Policy.ReadWrite.ConditionalAccess:
# confirmed against a live tenant, Get-MgIdentityConditionalAccessPolicy (the read
# cmdlet - called even from Set- to look up an existing toolkit-owned policy before
# deciding create vs. update) fails with "[AccessDenied]: required scopes are
# missing in the token" under Policy.ReadWrite.ConditionalAccess alone, even with
# that scope fully admin-consented and present on the token. Microsoft's own
# documented permission for this specific cmdlet is Policy.Read.All - Conditional
# Access does not follow the usual "ReadWrite implies Read" pattern other Graph
# resources do.
$script:GraphScopes = @(
    'Policy.ReadWrite.Authorization'
    'Policy.ReadWrite.AuthenticationMethod'
    'Directory.Read.All'
    'RoleManagement.Read.Directory'
    'Organization.Read.All'
    'Policy.Read.All'
    'Policy.ReadWrite.ConditionalAccess'
    'Group.ReadWrite.All'
    'Application.Read.All'
)

# A control's "workload" label in the config is the audit grouping used in reports.
# Its *connection* requirement (which service must be connected before Get-/Set- runs)
# is usually the same, but EntraID-UnifiedAuditLog is a documented exception: the
# cmdlet Microsoft ships for unified audit log ingestion is an Exchange Online
# cmdlet even though the setting is conceptually an EntraID/tenant-wide control.
# This map is orchestration metadata (how to connect), not desired-state data, so
# it stays in code rather than in the JSON config.
$script:ControlConnectionOverrides = @{
    'EntraID-UnifiedAuditLog' = 'ExchangeOnline'
}

# Controls whose primary Connection (above) isn't the only one they need. Currently
# just ExchangeOnline-AntiPhishingMailboxIntelligence: its actual Set-AntiPhishPolicy
# call is Exchange-only, but it also has to call Test-TenantServicePlan (Graph) to
# check for Defender for Office 365 licensing before doing anything. Every
# Conditional Access control's primary Connection is already 'Graph', so none of
# them need an entry here despite also calling Test-TenantServicePlan.
$script:ControlExtraConnections = @{
    'ExchangeOnline-AntiPhishingMailboxIntelligence' = @('Graph')
}

$script:WorkloadToConnectionDefault = @{
    EntraID           = 'Graph'
    ExchangeOnline    = 'ExchangeOnline'
    Teams             = 'Teams'
    SharePointOnline  = 'SharePointOnline'
    ConditionalAccess = 'Graph'
}

# Recognized Set- function result statuses.
$script:ResultStatus = @{
    Success              = 'Success'
    Failed               = 'Failed'
    SkippedAlreadyOk     = 'Skipped-AlreadyCompliant'
    SkippedManual        = 'Skipped-Manual'
    SkippedDisabled      = 'Skipped-Disabled'
}

# ---------------------------------------------------------------------------
# Config loading and validation
# ---------------------------------------------------------------------------

function Import-BaselineConfig {
    <#
    .SYNOPSIS
        Loads and validates the baseline desired-state config file.
    .DESCRIPTION
        Reads the JSON config file, validates it against the JSON schema (structural
        validation) and against a set of hand-rolled semantic rules (clear,
        control-specific error messages). Throws a single aggregated error listing
        every problem found rather than stopping at the first one. Never evaluates
        any string in the config file as code - every field is treated as inert data.
    .PARAMETER Path
        Path to the baseline.config.json file.
    .PARAMETER SchemaPath
        Path to the baseline.config.schema.json file.
    .EXAMPLE
        Import-BaselineConfig -Path ./config/baseline.config.json -SchemaPath ./config/baseline.config.schema.json
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$SchemaPath
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Baseline config file not found: $Path"
    }
    if (-not (Test-Path -LiteralPath $SchemaPath)) {
        throw "Baseline config schema file not found: $SchemaPath"
    }

    $rawJson = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop

    try {
        $null = $rawJson | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        throw "Baseline config file '$Path' is not valid JSON: $($_.Exception.Message)"
    }

    $schemaJson = Get-Content -LiteralPath $SchemaPath -Raw -ErrorAction Stop
    $schemaErrors = @()
    try {
        $null = Test-Json -Json $rawJson -Schema $schemaJson -ErrorAction Stop
    }
    catch {
        $schemaErrors += $_.Exception.Message
    }

    $config = $rawJson | ConvertFrom-Json -Depth 25

    if ($config.schemaVersion -ne $script:SupportedConfigSchemaVersion) {
        $schemaErrors += "Config schemaVersion '$($config.schemaVersion)' is not supported by this build of the toolkit (expected '$script:SupportedConfigSchemaVersion')."
    }

    $semanticErrors = Test-BaselineConfigSemantics -Config $config

    $allErrors = @($schemaErrors) + @($semanticErrors)
    if ($allErrors.Count -gt 0) {
        $message = "Baseline config validation failed with $($allErrors.Count) issue(s):`n" + (($allErrors | ForEach-Object { " - $_" }) -join "`n")
        throw $message
    }

    return $config
}

function Test-BaselineConfigSemantics {
    <#
    .SYNOPSIS
        Hand-rolled semantic validation of a parsed baseline config, beyond what
        the JSON schema alone can express.
    .DESCRIPTION
        Returns an array of human-readable error strings (empty array = valid).
        Checks: duplicate ids, automatable=false controls missing manualInstructions,
        Range compliance mode requiring a {min,max} desiredValue, and desiredValue
        typing sanity for the object-shaped controls in this toolkit's inventory.
    .PARAMETER Config
        The parsed config object (from ConvertFrom-Json).
    .EXAMPLE
        Test-BaselineConfigSemantics -Config $config
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [object]$Config
    )

    $errors = [System.Collections.Generic.List[string]]::new()

    if (-not $Config.PSObject.Properties['controls'] -or -not $Config.controls) {
        $errors.Add("Config has no 'controls' array.")
        return ,$errors.ToArray()
    }

    $seenIds = [System.Collections.Generic.HashSet[string]]::new()
    foreach ($control in $Config.controls) {
        $id = [string]$control.id
        if ([string]::IsNullOrWhiteSpace($id)) {
            $errors.Add("A control entry is missing an 'id'.")
            continue
        }
        if (-not $seenIds.Add($id)) {
            $errors.Add("Control '$id': duplicate id - each control id must be unique.")
        }

        if ($control.automatable -eq $false) {
            $manualInstructions = if ($control.PSObject.Properties['manualInstructions']) { [string]$control.manualInstructions } else { '' }
            if ([string]::IsNullOrWhiteSpace($manualInstructions)) {
                $errors.Add("Control '$id': automatable is false but 'manualInstructions' is missing/empty; the report cannot tell an admin where to change this by hand.")
            }
        }

        $complianceMode = if ($control.PSObject.Properties['complianceMode']) { [string]$control.complianceMode } else { 'Equality' }
        if ($complianceMode -eq 'Range') {
            $desired = $control.desiredValue
            $hasMin = $desired -and $desired.PSObject.Properties['min']
            $hasMax = $desired -and $desired.PSObject.Properties['max']
            if (-not ($hasMin -and $hasMax)) {
                $errors.Add("Control '$id': complianceMode is 'Range' but desiredValue is not a {min,max} object.")
            }
            elseif ([double]$desired.min -gt [double]$desired.max) {
                $errors.Add("Control '$id': desiredValue.min ($($desired.min)) is greater than desiredValue.max ($($desired.max)).")
            }
        }

        if ($control.PSObject.Properties['requiresPopulatedFields']) {
            foreach ($field in $control.requiresPopulatedFields) {
                if (-not ($control.desiredValue -and $control.desiredValue.PSObject.Properties[$field])) {
                    $errors.Add("Control '$id': requiresPopulatedFields references '$field', but desiredValue has no such property.")
                }
            }
        }
    }

    return ,$errors.ToArray()
}

function Get-BaselineControlConnection {
    <#
    .SYNOPSIS
        Resolves which backend connection a control needs (Graph, ExchangeOnline,
        Teams, or SharePointOnline), which may differ from its report 'workload'.
    .PARAMETER Control
        A single control entry from the parsed config.
    .EXAMPLE
        Get-BaselineControlConnection -Control $control
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [object]$Control
    )

    if ($script:ControlConnectionOverrides.ContainsKey($Control.id)) {
        return $script:ControlConnectionOverrides[$Control.id]
    }
    if ($script:WorkloadToConnectionDefault.ContainsKey($Control.workload)) {
        return $script:WorkloadToConnectionDefault[$Control.workload]
    }
    throw "Control '$($Control.id)': unable to resolve a required connection for workload '$($Control.workload)'."
}

function Get-BaselineControlExtraConnections {
    <#
    .SYNOPSIS
        Resolves any *additional* backend connections a control needs beyond its
        primary Get-BaselineControlConnection result (e.g. a control whose actual
        Set- call is Exchange-only but also needs Graph for a license check).
    .PARAMETER Control
        A single control entry from the parsed config.
    .EXAMPLE
        Get-BaselineControlExtraConnections -Control $control
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [object]$Control
    )
    if ($script:ControlExtraConnections.ContainsKey($Control.id)) {
        return ,[string[]]$script:ControlExtraConnections[$Control.id]
    }
    return ,[string[]]@()
}

# ---------------------------------------------------------------------------
# Control catalog
# ---------------------------------------------------------------------------

function Get-BaselineControlCatalog {
    <#
    .SYNOPSIS
        Builds the control catalog by matching every enabled config entry to its
        Get-<Id>State / Set-<Id>State function pair.
    .DESCRIPTION
        The orchestrator never invents behavior for a control it doesn't recognize:
        if a config entry's id has no matching Get-/Set- function pair loaded into
        the session, or a control module exposes a Get-/Set- pair with no matching
        config entry, that is a validation error surfaced before any connection is
        made or any state is read/changed.
    .PARAMETER Config
        The parsed, already-validated config object.
    .PARAMETER AvailableFunctions
        Optional override of the function name list to match against (for testing).
        Defaults to every Get-*State/Set-*State command currently loaded.
    .EXAMPLE
        Get-BaselineControlCatalog -Config $config
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param(
        [Parameter(Mandatory)]
        [object]$Config,

        [Parameter()]
        [string[]]$AvailableFunctions
    )

    if (-not $AvailableFunctions) {
        # Scoped to our own four control modules by name, not just the Get-/Set-*State
        # naming pattern: some vendor modules ship real cmdlets that happen to match
        # it too - e.g. Microsoft.Online.SharePoint.PowerShell's genuine
        # Get-/Set-SPOStructuralNavigationCacheSiteState and ...CacheWebState, which
        # otherwise get misidentified as orphaned catalog controls once that module is
        # loaded (it stays loaded across separate runs of this script in the same
        # PowerShell window when -KeepConnectionsOpen was used on a prior run).
        $ourModuleNames = @('EntraIdControls', 'ExchangeOnlineControls', 'TeamsControls', 'SharePointOnlineControls', 'ConditionalAccessControls')
        $AvailableFunctions = (Get-Command -CommandType Function | Where-Object { $_.Name -match '^(Get|Set)-.+State$' -and $_.ModuleName -in $ourModuleNames }).Name
    }

    $errors = [System.Collections.Generic.List[string]]::new()
    $catalog = [System.Collections.Generic.List[pscustomobject]]::new()
    $matchedFunctionNames = [System.Collections.Generic.HashSet[string]]::new()

    $enabledControls = @($Config.controls | Where-Object { $_.enabled })

    foreach ($control in $enabledControls) {
        $getName = "Get-$($control.id)State"
        $setName = "Set-$($control.id)State"

        $hasGet = $AvailableFunctions -contains $getName
        $hasSet = $AvailableFunctions -contains $setName

        if (-not $hasGet) { $errors.Add("Control '$($control.id)' is enabled in config but no '$getName' function is implemented in the control catalog.") }
        if (-not $hasSet) { $errors.Add("Control '$($control.id)' is enabled in config but no '$setName' function is implemented in the control catalog.") }

        if ($hasGet -and $hasSet) {
            [void]$matchedFunctionNames.Add($getName)
            [void]$matchedFunctionNames.Add($setName)
            $catalog.Add([pscustomobject]@{
                Id           = $control.id
                Workload     = $control.workload
                Connection   = Get-BaselineControlConnection -Control $control
                ExtraConnections = Get-BaselineControlExtraConnections -Control $control
                Automatable  = [bool]$control.automatable
                DesiredValue = $control.desiredValue
                Description  = $control.description
                ComplianceMode = if ($control.PSObject.Properties['complianceMode']) { [string]$control.complianceMode } else { 'Equality' }
                ManualInstructions = if ($control.PSObject.Properties['manualInstructions']) { [string]$control.manualInstructions } else { '' }
                RequiresPopulatedFields = @(if ($control.PSObject.Properties['requiresPopulatedFields']) { $control.requiresPopulatedFields } else { @() })
                Tier         = if ($control.PSObject.Properties['tier']) { [Nullable[int]][int]$control.tier } else { $null }
                ForceCreateDespiteOverlap = if ($control.PSObject.Properties['forceCreateDespiteOverlap']) { [bool]$control.forceCreateDespiteOverlap } else { $false }
                GetCommand   = $getName
                SetCommand   = $setName
            })
        }
    }

    # Orphaned catalog functions: implemented but not referenced by any enabled config entry.
    $configuredIds = [System.Collections.Generic.HashSet[string]]::new([string[]]$enabledControls.id)
    foreach ($fn in $AvailableFunctions) {
        if ($fn -match '^Get-(.+)State$') {
            $candidateId = $Matches[1]
            if (-not $configuredIds.Contains($candidateId) -and ($AvailableFunctions -contains "Set-${candidateId}State")) {
                # Only flag it if the id isn't present at all in config (vs. present-but-disabled, which is a legitimate opt-out).
                $presentButDisabled = @($Config.controls | Where-Object { $_.id -eq $candidateId }).Count -gt 0
                if (-not $presentButDisabled) {
                    $errors.Add("Control catalog implements '$candidateId' (Get-${candidateId}State/Set-${candidateId}State) but config/baseline.config.json has no entry for it.")
                }
            }
        }
    }

    if ($errors.Count -gt 0) {
        $message = "Control catalog validation failed with $($errors.Count) issue(s):`n" + (($errors | ForEach-Object { " - $_" }) -join "`n")
        throw $message
    }

    return ,$catalog.ToArray()
}

# ---------------------------------------------------------------------------
# Value comparison
# ---------------------------------------------------------------------------

function Compare-BaselineValueDeep {
    <#
    .SYNOPSIS
        Deep, order-insensitive-for-objects structural equality check used to
        compute compliance between a current value and a desired value.
    .PARAMETER Left
        First value (typically the live/current value).
    .PARAMETER Right
        Second value (typically the desired value from config).
    .EXAMPLE
        Compare-BaselineValueDeep -Left $current -Right $desired
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Left,

        [Parameter()]
        [AllowNull()]
        [object]$Right
    )

    if ($null -eq $Left -and $null -eq $Right) { return $true }
    if ($null -eq $Left -or $null -eq $Right) { return $false }

    $leftIsCollection = ($Left -is [System.Collections.IEnumerable]) -and (-not ($Left -is [string]))
    $rightIsCollection = ($Right -is [System.Collections.IEnumerable]) -and (-not ($Right -is [string]))

    if ($leftIsCollection -and $rightIsCollection) {
        $leftArr = @($Left)
        $rightArr = @($Right)
        if ($leftArr.Count -ne $rightArr.Count) { return $false }
        for ($i = 0; $i -lt $leftArr.Count; $i++) {
            if (-not (Compare-BaselineValueDeep -Left $leftArr[$i] -Right $rightArr[$i])) { return $false }
        }
        return $true
    }
    if ($leftIsCollection -ne $rightIsCollection) { return $false }

    $leftIsObject = ($Left -is [System.Management.Automation.PSCustomObject]) -or ($Left -is [System.Collections.IDictionary])
    $rightIsObject = ($Right -is [System.Management.Automation.PSCustomObject]) -or ($Right -is [System.Collections.IDictionary])

    if ($leftIsObject -and $rightIsObject) {
        $leftProps = Get-BaselinePropertyMap -Value $Left
        $rightProps = Get-BaselinePropertyMap -Value $Right
        $leftKeys = [string[]]$leftProps.Keys
        $rightKeys = [string[]]$rightProps.Keys
        if (Compare-Object -ReferenceObject $leftKeys -DifferenceObject $rightKeys -CaseSensitive:$false) { return $false }
        foreach ($key in $leftKeys) {
            $matchKey = $rightProps.Keys | Where-Object { $_ -ieq $key } | Select-Object -First 1
            if (-not (Compare-BaselineValueDeep -Left $leftProps[$key] -Right $rightProps[$matchKey])) { return $false }
        }
        return $true
    }
    if ($leftIsObject -ne $rightIsObject) { return $false }

    # Scalars: compare loosely on numerics (int vs double from JSON), strictly otherwise.
    if (($Left -is [ValueType] -or $Left -is [string]) -and ($Right -is [ValueType] -or $Right -is [string])) {
        if ($Left -is [bool] -or $Right -is [bool]) {
            return ([bool]$Left) -eq ([bool]$Right)
        }
        if (($Left -is [string]) -or ($Right -is [string])) {
            return [string]$Left -eq [string]$Right
        }
        try {
            return ([double]$Left) -eq ([double]$Right)
        }
        catch {
            return $Left -eq $Right
        }
    }

    return $Left -eq $Right
}

function Get-BaselinePropertyMap {
    <#
    .SYNOPSIS
        Normalizes a PSCustomObject or hashtable/dictionary into an ordered
        hashtable of property name/value pairs for comparison and templating.
    .PARAMETER Value
        The object to normalize.
    .EXAMPLE
        Get-BaselinePropertyMap -Value $desiredValue
    #>
    [CmdletBinding()]
    [OutputType([hashtable])]
    param(
        [Parameter(Mandatory)]
        [object]$Value
    )

    $map = [ordered]@{}
    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($key in $Value.Keys) { $map[[string]$key] = $Value[$key] }
    }
    else {
        foreach ($p in $Value.PSObject.Properties) { $map[$p.Name] = $p.Value }
    }
    return $map
}

function Test-BaselineCompliance {
    <#
    .SYNOPSIS
        Computes whether a current value is compliant with a desired value under
        a given compliance mode.
    .PARAMETER CurrentValue
        The live value read from the tenant.
    .PARAMETER DesiredValue
        The value from config.
    .PARAMETER ComplianceMode
        'Equality' (default) does a deep structural comparison. 'Range' expects
        DesiredValue to be a {min,max} object and CurrentValue to be numeric.
    .EXAMPLE
        Test-BaselineCompliance -CurrentValue 3 -DesiredValue @{min=2;max=4} -ComplianceMode Range
    #>
    [CmdletBinding()]
    [OutputType([Nullable[bool]])]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [ValidateSet('Equality', 'Range')]
        [string]$ComplianceMode = 'Equality'
    )

    if ($null -eq $CurrentValue) { return $null }

    if ($ComplianceMode -eq 'Range') {
        $props = Get-BaselinePropertyMap -Value $DesiredValue
        return ([double]$CurrentValue -ge [double]$props['min']) -and ([double]$CurrentValue -le [double]$props['max'])
    }

    return Compare-BaselineValueDeep -Left $CurrentValue -Right $DesiredValue
}

# ---------------------------------------------------------------------------
# Licensing
# ---------------------------------------------------------------------------

# Module-scoped, not session-global: Invoke-M365Baseline.ps1 re-imports this module
# with -Force on every run, which resets this back to $null, so a stale license
# read from an earlier run in the same PowerShell window can never leak into a
# later one. Within a single run it's populated once and reused by every caller.
$script:SubscribedSkuCache = $null

function Get-BaselineSubscribedSkuCache {
    <#
    .SYNOPSIS
        Internal: returns this run's cached Get-MgSubscribedSku result, calling it
        only on first use (or when -Refresh is passed) and reusing the result for
        every subsequent license check in the same run.
    .DESCRIPTION
        Backs Test-TenantServicePlan. Both Conditional Access license gating and
        ExchangeOnline-AntiPhishingMailboxIntelligence's Defender for Office 365
        gate call Test-TenantServicePlan, potentially many times across many
        controls in one run - this cache is what keeps Get-MgSubscribedSku itself
        to at most one real call per run.
    .PARAMETER Refresh
        Forces a fresh Get-MgSubscribedSku call even if a cached result exists.
        Not used by Test-TenantServicePlan itself; available for callers (e.g.
        tests) that need to invalidate the cache mid-run.
    .EXAMPLE
        Get-BaselineSubscribedSkuCache
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter()]
        [switch]$Refresh
    )
    if ($Refresh -or $null -eq $script:SubscribedSkuCache) {
        $script:SubscribedSkuCache = @(Get-MgSubscribedSku -All -ErrorAction Stop)
    }
    return ,$script:SubscribedSkuCache
}

function Test-TenantServicePlan {
    <#
    .SYNOPSIS
        Checks whether the tenant holds at least one of the given Microsoft Graph
        service plan names, in an active provisioning state, across all of its
        subscribed SKUs.
    .DESCRIPTION
        Shared license-gate infrastructure - the only code path in this toolkit
        that calls Get-MgSubscribedSku (via Get-BaselineSubscribedSkuCache, which
        caches the result for the rest of the run). Used by both
        ConditionalAccessControls.psm1 (Entra ID P1/P2 gating) and
        ExchangeOnline-AntiPhishingMailboxIntelligence (Defender for Office 365
        Plan 1/2 gating) - any future license-gated control should reuse this
        rather than calling Get-MgSubscribedSku directly.

        A service plan name is matched against every subscribed SKU's ServicePlans
        collection, not just one specific SKU, since the same service plan (e.g.
        AAD_PREMIUM) can be granted by more than one SKU. "Active" means any
        provisioning status other than 'Disabled': an admin can disable one
        service plan within an otherwise-active SKU without removing the SKU
        itself, but 'PendingActivation'/'PendingInput'/'PendingProvisioning' all
        still mean the plan is granted (just not fully rolled out yet), not absent.

        Requires a Microsoft Graph connection with at least Organization.Read.All.
    .PARAMETER ServicePlanNames
        One or more Graph service plan names (e.g. 'AAD_PREMIUM'). Returns $true
        if ANY of them is present and active - callers that need "either of these
        two plans satisfies the gate" (e.g. ATP_ENTERPRISE or THREAT_INTELLIGENCE
        for Defender for Office 365 Plan 1 or 2) pass both in one call.
    .EXAMPLE
        Test-TenantServicePlan -ServicePlanNames @('AAD_PREMIUM')
    .EXAMPLE
        Test-TenantServicePlan -ServicePlanNames @('ATP_ENTERPRISE','THREAT_INTELLIGENCE')
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string[]]$ServicePlanNames
    )
    $skus = Get-BaselineSubscribedSkuCache
    foreach ($sku in $skus) {
        foreach ($plan in @($sku.ServicePlans)) {
            if (($ServicePlanNames -contains $plan.ServicePlanName) -and ([string]$plan.ProvisioningStatus -ne 'Disabled')) {
                return $true
            }
        }
    }
    return $false
}

# ---------------------------------------------------------------------------
# Module / connection management
# ---------------------------------------------------------------------------

function Test-BaselineRequiredModule {
    <#
    .SYNOPSIS
        Checks whether a required PowerShell module is installed.
    .DESCRIPTION
        Microsoft.Online.SharePoint.PowerShell is a special case: it's only ever
        loaded inside a Windows PowerShell 5.1 compatibility session (via
        Connect-BaselineWorkload's -UseWindowsPowerShell import), which has its
        own separate CurrentUser module path from PowerShell 7
        (Documents\WindowsPowerShell\Modules vs. Documents\PowerShell\Modules on
        Windows). Checking with Get-Module from inside PS7 would answer whether
        PS7 can see it, not whether that compatibility session can - so this one
        module is checked via a real Windows PowerShell 5.1 process instead.
    .PARAMETER Name
        Module name.
    .EXAMPLE
        Test-BaselineRequiredModule -Name ExchangeOnlineManagement
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )
    if ($Name -eq 'Microsoft.Online.SharePoint.PowerShell') {
        $output = & powershell.exe -NoProfile -Command "[bool](Get-Module -ListAvailable -Name '$Name')" 2>$null
        return ([string]$output).Trim() -eq 'True'
    }
    return [bool](Get-Module -ListAvailable -Name $Name | Select-Object -First 1)
}

function Install-BaselineRequiredModule {
    <#
    .SYNOPSIS
        Installs a required module from PSGallery for the current user.
    .DESCRIPTION
        Microsoft.Online.SharePoint.PowerShell is installed via a real Windows
        PowerShell 5.1 process rather than the current PowerShell 7 session -
        see Test-BaselineRequiredModule for why: PS7's Install-Module would put
        it somewhere the -UseWindowsPowerShell compatibility session that
        actually loads this module never looks.
    .PARAMETER Name
        Module name to install.
    .EXAMPLE
        Install-BaselineRequiredModule -Name MicrosoftTeams
    #>
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )
    if ($PSCmdlet.ShouldProcess($Name, 'Install-Module -Scope CurrentUser')) {
        if ($Name -eq 'Microsoft.Online.SharePoint.PowerShell') {
            & powershell.exe -NoProfile -Command "Install-Module -Name '$Name' -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop"
            if ($LASTEXITCODE -ne 0) {
                throw "Failed to install $Name via Windows PowerShell 5.1 (exit code $LASTEXITCODE)."
            }
            return
        }
        Install-Module -Name $Name -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
}

function Assert-BaselineRequiredModules {
    <#
    .SYNOPSIS
        Ensures every module needed by the given connections is installed,
        optionally auto-installing missing ones.
    .PARAMETER Connections
        Distinct connection names ('Graph','ExchangeOnline','Teams','SharePointOnline').
    .PARAMETER InstallMissingModules
        If set, missing modules are installed from PSGallery for CurrentUser.
        Otherwise a missing module is a terminating, actionable error.
    .EXAMPLE
        Assert-BaselineRequiredModules -Connections 'Graph','ExchangeOnline' -InstallMissingModules
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Connections,

        [Parameter()]
        [switch]$InstallMissingModules
    )

    foreach ($conn in ($Connections | Select-Object -Unique)) {
        $moduleName = $script:WorkloadModuleMap[$conn]
        if (-not $moduleName) { throw "Unknown connection type '$conn'." }

        if (-not (Test-BaselineRequiredModule -Name $moduleName)) {
            if ($InstallMissingModules) {
                Write-Verbose "Installing missing module '$moduleName' (Scope CurrentUser)..."
                Install-BaselineRequiredModule -Name $moduleName
            }
            else {
                throw "Required module '$moduleName' (for $conn) is not installed. Install it with: Install-Module -Name $moduleName -Scope CurrentUser, or re-run with -InstallMissingModules."
            }
        }
    }
}

function Get-BaselineConnectionOrder {
    <#
    .SYNOPSIS
        Orders a list of required connections into the sequence they should be
        connected in.
    .DESCRIPTION
        Returns the subset of $script:WorkloadModuleMap's canonical key order
        (Graph, ExchangeOnline, Teams, SharePointOnline) that appears in
        -Connections. Graph must connect before ExchangeOnline in the same
        process to avoid a known MSAL/Microsoft.Identity.Client assembly
        version conflict between the Microsoft.Graph and ExchangeOnlineManagement
        modules (see the comment on $script:WorkloadModuleMap) - callers should
        always connect in this order rather than in config/catalog-encounter order.
    .PARAMETER Connections
        Connection names to order (any of 'Graph','ExchangeOnline','Teams','SharePointOnline').
    .EXAMPLE
        Get-BaselineConnectionOrder -Connections 'ExchangeOnline','Graph'
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [string[]]$Connections
    )
    $distinct = $Connections | Select-Object -Unique
    return ,[string[]]($script:WorkloadModuleMap.Keys | Where-Object { $distinct -contains $_ })
}

function Test-BaselineWorkloadConnected {
    <#
    .SYNOPSIS
        Checks whether a connection was left open by a prior run in this same
        PowerShell session (via -KeepConnectionsOpen).
    .DESCRIPTION
        Connect-MgGraph, Connect-ExchangeOnline, Connect-MicrosoftTeams, and
        Connect-SPOService do not check for an existing session themselves -
        each one unconditionally starts a fresh interactive sign-in whenever
        it's called, live session or not. So merely skipping the disconnect
        step (-KeepConnectionsOpen) does nothing on its own; the caller also
        has to skip re-calling Connect- for a connection that's already live.
        Tracking state lives in a true PowerShell session global variable
        (not a module-scoped one) because Invoke-M365Baseline.ps1 re-imports
        this module with -Force on every run, which would otherwise reset
        module-scoped state and defeat the whole point.
    .PARAMETER Connection
        'Graph', 'ExchangeOnline', 'Teams', or 'SharePointOnline'.
    .EXAMPLE
        Test-BaselineWorkloadConnected -Connection Graph
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'Teams', 'SharePointOnline')]
        [string]$Connection
    )
    if (-not (Test-Path Variable:Global:M365BaselineActiveConnections)) { return $false }
    return [bool]$Global:M365BaselineActiveConnections.Contains($Connection)
}

function Set-BaselineWorkloadConnectedState {
    <#
    .SYNOPSIS
        Internal: records that a connection is (or is no longer) live in the
        session-global tracking set used by Test-BaselineWorkloadConnected.
    .PARAMETER Connection
        'Graph', 'ExchangeOnline', 'Teams', or 'SharePointOnline'.
    .PARAMETER Connected
        $true to mark it live, $false to clear it.
    .EXAMPLE
        Set-BaselineWorkloadConnectedState -Connection Graph -Connected $true
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'Teams', 'SharePointOnline')]
        [string]$Connection,

        [Parameter(Mandatory)]
        [bool]$Connected
    )
    if (-not (Test-Path Variable:Global:M365BaselineActiveConnections)) {
        $Global:M365BaselineActiveConnections = [System.Collections.Generic.HashSet[string]]::new()
    }
    if ($Connected) { [void]$Global:M365BaselineActiveConnections.Add($Connection) }
    else { [void]$Global:M365BaselineActiveConnections.Remove($Connection) }
}

function Connect-BaselineWorkload {
    <#
    .SYNOPSIS
        Establishes a connection for a single backend service, once per run.
    .DESCRIPTION
        Always performs a real connection attempt - callers that want to reuse
        a connection left open by a prior run (-KeepConnectionsOpen) should
        check Test-BaselineWorkloadConnected first and skip calling this at
        all when it returns $true, since Connect-MgGraph/Connect-ExchangeOnline/
        Connect-MicrosoftTeams/Connect-SPOService each start a fresh interactive
        sign-in unconditionally rather than detecting an existing session.
    .PARAMETER Connection
        'Graph', 'ExchangeOnline', 'Teams', or 'SharePointOnline'.
    .PARAMETER SharePointAdminUrl
        Required only when Connection is 'SharePointOnline' (e.g. https://contoso-admin.sharepoint.com).
    .EXAMPLE
        Connect-BaselineWorkload -Connection Graph
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'Teams', 'SharePointOnline')]
        [string]$Connection,

        [Parameter()]
        [string]$SharePointAdminUrl
    )

    try {
        switch ($Connection) {
            'Graph' {
                Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
                Connect-MgGraph -Scopes $script:GraphScopes -NoWelcome -ErrorAction Stop
            }
            'ExchangeOnline' {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                # ExchangeOnlineManagement 3.7+ enables Windows Account Manager (WAM)
                # sign-in by default. WAM occasionally crashes with a NullReferenceException
                # in Microsoft.Identity.Client's RuntimeBroker when another module (e.g.
                # Microsoft.Graph) has already used MSAL earlier in the same process - see
                # https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3576.
                # -DisableWAM avoids that crash by falling back to the older, broker-free
                # interactive browser flow. It should not be forced unconditionally though,
                # since it's a strictly worse flow when WAM isn't actually failing: attempt
                # the normal WAM connection first, and fall back to -DisableWAM only if that
                # specific RuntimeBroker crash actually happens.
                #
                # (An earlier version of this comment blamed -DisableWAM for HTTP 403s on
                # Get-ExternalInOutlook/ExchangeOnline-ExternalSenderTag. That was wrong -
                # the real cause was unrelated, in how that control called
                # Get-ExternalInOutlook; see Get-ExchangeOnline-ExternalSenderTagState in
                # ExchangeOnlineControls.psm1 for the actual root cause and fix.)
                $eopParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
                $supportsDisableWam = (Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')
                try {
                    Connect-ExchangeOnline @eopParams
                }
                catch {
                    $isWamBrokerCrash = $supportsDisableWam -and
                        ($_.Exception -is [System.NullReferenceException] -or
                         $_.Exception.ToString() -match 'RuntimeBroker' -or
                         $_.ToString() -match 'RuntimeBroker')
                    if (-not $isWamBrokerCrash) { throw }

                    Write-Warning "Connect-ExchangeOnline failed with what looks like the known WAM/RuntimeBroker crash; retrying with -DisableWAM."
                    $eopParams['DisableWAM'] = $true
                    Connect-ExchangeOnline @eopParams
                }
            }
            'Teams' {
                Import-Module MicrosoftTeams -ErrorAction Stop
                Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
            }
            'SharePointOnline' {
                if ([string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
                    throw "SharePointOnline connection requires -SharePointAdminUrl (e.g. https://contoso-admin.sharepoint.com)."
                }
                # Microsoft.Online.SharePoint.PowerShell targets .NET Framework, not
                # PowerShell 7's .NET runtime, and its own OAuth/broker handling is
                # unreliable when loaded directly into a PS7 process that has already
                # used MSAL for Graph/Exchange/Teams (surfaces as "No valid OAuth 2.0
                # authentication session exists" from Connect-SPOService even with a
                # correct URL and role). -UseWindowsPowerShell loads the module into an
                # isolated background Windows PowerShell 5.1 process via implicit
                # remoting - the documented workaround for this module on PS7+ - which
                # also sidesteps any broker state left over from the other connections.
                # -Global is required here specifically: Import-Module normally adds a
                # module's commands to the session for every caller regardless of scope,
                # but -UseWindowsPowerShell instead dynamically generates local proxy
                # functions for the implicitly-remoted commands, and without -Global
                # those proxies are only visible inside this function's own scope - the
                # control functions in SharePointOnlineControls.psm1 (a sibling module)
                # would otherwise see "Get-SPOTenant is not recognized" even though this
                # Connect-SPOService call two lines down succeeds.
                Import-Module Microsoft.Online.SharePoint.PowerShell -UseWindowsPowerShell -Global -ErrorAction Stop
                Connect-SPOService -Url $SharePointAdminUrl -ErrorAction Stop
            }
        }
    }
    catch {
        throw "Failed to connect to $Connection`: $($_.Exception.Message). Verify the account has the admin role required for this workload (see README.md)."
    }

    Set-BaselineWorkloadConnectedState -Connection $Connection -Connected $true
}

function Disconnect-BaselineWorkload {
    <#
    .SYNOPSIS
        Best-effort disconnect for a single backend service. Never throws.
    .PARAMETER Connection
        'Graph', 'ExchangeOnline', 'Teams', or 'SharePointOnline'.
    .EXAMPLE
        Disconnect-BaselineWorkload -Connection Graph
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet('Graph', 'ExchangeOnline', 'Teams', 'SharePointOnline')]
        [string]$Connection
    )
    try {
        switch ($Connection) {
            'Graph' { Disconnect-MgGraph -ErrorAction Stop | Out-Null }
            'ExchangeOnline' { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop }
            'Teams' { Disconnect-MicrosoftTeams -ErrorAction Stop }
            'SharePointOnline' { Disconnect-SPOService -ErrorAction Stop }
        }
    }
    catch {
        Write-Verbose "Non-fatal: disconnect from $Connection reported: $($_.Exception.Message)"
    }
    finally {
        # Clear the tracking flag even if the disconnect call itself failed:
        # better to reconnect fresh next run than to treat a possibly-broken
        # session as reusable.
        Set-BaselineWorkloadConnectedState -Connection $Connection -Connected $false
    }
}

# ---------------------------------------------------------------------------
# Audit engine
# ---------------------------------------------------------------------------

function Invoke-BaselineControlAudit {
    <#
    .SYNOPSIS
        Reads current state for every control in the catalog and computes compliance.
    .DESCRIPTION
        Calls each control's Get-<Id>State function. A control that throws while
        being read is recorded with Compliant = $null and its error message,
        contributing to the audit's error count, which is distinct from a normal
        non-compliant finding.
    .PARAMETER Catalog
        Catalog entries from Get-BaselineControlCatalog.
    .EXAMPLE
        Invoke-BaselineControlAudit -Catalog $catalog
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param(
        [Parameter(Mandatory)]
        [pscustomobject[]]$Catalog
    )

    $results = [System.Collections.Generic.List[pscustomobject]]::new()

    foreach ($entry in $Catalog) {
        $currentValue = $null
        $errorMessage = $null
        try {
            $stateResult = & $entry.GetCommand
            $currentValue = $stateResult.Value
        }
        catch {
            $errorMessage = $_.Exception.Message
        }

        $compliant = if ($errorMessage) { $null } else { Test-BaselineCompliance -CurrentValue $currentValue -DesiredValue $entry.DesiredValue -ComplianceMode $entry.ComplianceMode }

        $results.Add([pscustomobject]@{
            Id                 = $entry.Id
            Workload           = $entry.Workload
            Description        = $entry.Description
            Automatable        = $entry.Automatable
            CurrentValue       = $currentValue
            DesiredValue       = $entry.DesiredValue
            Compliant          = $compliant
            ManualInstructions = $entry.ManualInstructions
            Error              = $errorMessage
        })
    }

    return ,$results.ToArray()
}

# ---------------------------------------------------------------------------
# Apply / Restore engine
# ---------------------------------------------------------------------------

function Test-BaselineApplyReadiness {
    <#
    .SYNOPSIS
        Pre-flight check for Apply mode: catches controls whose desiredValue is
        missing tenant-specific required fields (e.g. empty domain lists) before
        any connection is made or any change is attempted.
    .PARAMETER Catalog
        Catalog entries from Get-BaselineControlCatalog.
    .EXAMPLE
        Test-BaselineApplyReadiness -Catalog $catalog
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory)]
        [pscustomobject[]]$Catalog
    )

    $errors = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $Catalog) {
        if (-not $entry.Automatable) { continue }
        foreach ($field in $entry.RequiresPopulatedFields) {
            $value = $null
            if ($entry.DesiredValue -and $entry.DesiredValue.PSObject.Properties[$field]) {
                $value = $entry.DesiredValue.$field
            }
            $isEmpty = ($null -eq $value) -or (($value -is [System.Collections.IEnumerable]) -and (-not ($value -is [string])) -and (@($value).Count -eq 0)) -or ($value -eq '')
            if ($isEmpty) {
                $errors.Add("Control '$($entry.Id)': desiredValue.$field is empty. This control requires tenant-specific values before it can be applied - edit config/baseline.config.json and populate it, or disable the control for this run.")
            }
        }
    }
    return ,$errors.ToArray()
}

function Invoke-BaselineControlApply {
    <#
    .SYNOPSIS
        Applies desired state for every enabled, automatable, non-compliant control.
    .DESCRIPTION
        Already-compliant controls are skipped and logged as Skipped-AlreadyCompliant
        without calling Set-. Non-automatable controls are skipped and logged as
        Skipped-Manual using the message their Set- function (or config
        manualInstructions) provides. Every attempt is appended to the change log.
        Honors ShouldProcess: when WhatIfMode is set, no Set- function is called.
    .PARAMETER Catalog
        Catalog entries from Get-BaselineControlCatalog.
    .PARAMETER AuditResults
        Pre-change audit results (from Invoke-BaselineControlAudit) used to decide
        what needs changing and to skip controls that errored on read.
    .PARAMETER ChangeLogPath
        Path to the JSON Lines change log file to append to.
    .PARAMETER WhatIfMode
        When set, simulates the run (calls ShouldProcess but never invokes Set-).
    .PARAMETER ShouldProcessTarget
        The cmdlet/script whose ShouldProcess gate to honor (pass $PSCmdlet from the caller).
    .PARAMETER AcknowledgeRisk
        Forwarded to Set- functions that accept an -AcknowledgeRisk switch, for
        controls with a deliberately severe empty-list interpretation.
    .EXAMPLE
        Invoke-BaselineControlApply -Catalog $catalog -AuditResults $audit -ChangeLogPath $log -ShouldProcessTarget $PSCmdlet
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param(
        [Parameter(Mandatory)]
        [pscustomobject[]]$Catalog,

        [Parameter(Mandatory)]
        [pscustomobject[]]$AuditResults,

        [Parameter(Mandatory)]
        [string]$ChangeLogPath,

        [Parameter()]
        [switch]$WhatIfMode,

        [Parameter()]
        [object]$ShouldProcessTarget,

        [Parameter()]
        [switch]$AcknowledgeRisk,

        [Parameter()]
        [switch]$StopOnError
    )

    $results = [System.Collections.Generic.List[pscustomobject]]::new()
    $auditById = @{}
    foreach ($a in $AuditResults) { $auditById[$a.Id] = $a }

    foreach ($entry in $Catalog) {
        $audit = $auditById[$entry.Id]
        if (-not $audit) { continue }

        $logRecord = [ordered]@{
            timestampUtc  = (Get-Date).ToUniversalTime().ToString('o')
            id            = $entry.Id
            workload      = $entry.Workload
            previousValue = $audit.CurrentValue
            attemptedValue = $entry.DesiredValue
            result        = $null
            errorMessage  = $null
        }

        if ($audit.Error) {
            $logRecord.result = $script:ResultStatus.Failed
            $logRecord.errorMessage = "Skipped: pre-change audit could not read current state ($($audit.Error))."
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = $script:ResultStatus.Failed; PreviousValue = $audit.CurrentValue; AppliedValue = $null; Message = $logRecord.errorMessage })
            if ($StopOnError) { throw "Stopping (StopOnError): $($entry.Id) - $($logRecord.errorMessage)" }
            continue
        }

        if (-not $entry.Automatable) {
            $message = if ($entry.ManualInstructions) { $entry.ManualInstructions } else { 'No automated remediation is implemented for this control.' }
            $logRecord.result = $script:ResultStatus.SkippedManual
            $logRecord.errorMessage = $message
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = $script:ResultStatus.SkippedManual; PreviousValue = $audit.CurrentValue; AppliedValue = $null; Message = $message })
            continue
        }

        if ($audit.Compliant -eq $true) {
            $logRecord.result = $script:ResultStatus.SkippedAlreadyOk
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = $script:ResultStatus.SkippedAlreadyOk; PreviousValue = $audit.CurrentValue; AppliedValue = $audit.CurrentValue; Message = 'Already compliant, no action.' })
            continue
        }

        $target = "$($entry.Id) ($($entry.Workload))"
        $action = "Set desired state (current: $($audit.CurrentValue | ConvertTo-Json -Compress -Depth 10) -> desired: $($entry.DesiredValue | ConvertTo-Json -Compress -Depth 10))"

        $shouldProceed = $true
        if ($ShouldProcessTarget -and ($ShouldProcessTarget | Get-Member -Name ShouldProcess -ErrorAction SilentlyContinue)) {
            $shouldProceed = $ShouldProcessTarget.ShouldProcess($target, $action)
        }

        if (-not $shouldProceed -or $WhatIfMode) {
            $logRecord.result = 'Skipped-WhatIf'
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = 'Skipped-WhatIf'; PreviousValue = $audit.CurrentValue; AppliedValue = $null; Message = 'Skipped due to -WhatIf.' })
            continue
        }

        try {
            $setParams = @{ DesiredValue = $entry.DesiredValue; CurrentValue = $audit.CurrentValue }
            $setCmd = Get-Command $entry.SetCommand -ErrorAction Stop
            if ($setCmd.Parameters.ContainsKey('AcknowledgeRisk')) {
                $setParams['AcknowledgeRisk'] = [bool]$AcknowledgeRisk
            }
            # PSObject.Properties[...] checks (not just .Tier/.ForceCreateDespiteOverlap
            # direct access) because catalog entries built by hand rather than via
            # Get-BaselineControlCatalog (e.g. Orchestrator.Tests.ps1's fake catalogs)
            # may not have these properties at all, and this module runs under
            # Set-StrictMode -Version Latest - referencing a genuinely absent property
            # throws rather than returning $null.
            if ($setCmd.Parameters.ContainsKey('Tier') -and $entry.PSObject.Properties['Tier'] -and $null -ne $entry.Tier) {
                $setParams['Tier'] = $entry.Tier
            }
            if ($setCmd.Parameters.ContainsKey('ForceCreateDespiteOverlap')) {
                $setParams['ForceCreateDespiteOverlap'] = if ($entry.PSObject.Properties['ForceCreateDespiteOverlap']) { [bool]$entry.ForceCreateDespiteOverlap } else { $false }
            }
            $setResult = & $entry.SetCommand @setParams

            $logRecord.result = $setResult.Status
            $logRecord.errorMessage = $setResult.Message
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add($setResult)
        }
        catch {
            $logRecord.result = $script:ResultStatus.Failed
            $logRecord.errorMessage = $_.Exception.Message
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = $script:ResultStatus.Failed; PreviousValue = $audit.CurrentValue; AppliedValue = $null; Message = $_.Exception.Message })
            if ($StopOnError) { throw "Stopping (StopOnError): $($entry.Id) - $($_.Exception.Message)" }
        }
    }

    return ,$results.ToArray()
}

function Invoke-BaselineControlRestore {
    <#
    .SYNOPSIS
        Replays a snapshot file's recorded current values as the new desired state.
    .DESCRIPTION
        For each control in the snapshot, calls the same Set-<Id>State function used
        by Apply, but passes the snapshot's currentValue (captured at snapshot time)
        as the value to converge to - never the live config's desiredValue.
    .PARAMETER Catalog
        Catalog entries from Get-BaselineControlCatalog (built from the *current* config,
        used only to resolve Set- function names / automatable flags).
    .PARAMETER SnapshotControls
        The 'controls' array from a loaded snapshot file.
    .PARAMETER ChangeLogPath
        Path to the JSON Lines change log file to append to.
    .PARAMETER WhatIfMode
        When set, simulates the run.
    .PARAMETER ShouldProcessTarget
        Pass $PSCmdlet from the caller to honor -WhatIf/-Confirm.
    .PARAMETER StopOnError
        Abort the whole restore on the first failure instead of continuing.
    .EXAMPLE
        Invoke-BaselineControlRestore -Catalog $catalog -SnapshotControls $snapshot.controls -ChangeLogPath $log -ShouldProcessTarget $PSCmdlet
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject[]])]
    param(
        [Parameter(Mandatory)]
        [pscustomobject[]]$Catalog,

        [Parameter(Mandatory)]
        [object[]]$SnapshotControls,

        [Parameter(Mandatory)]
        [string]$ChangeLogPath,

        [Parameter()]
        [switch]$WhatIfMode,

        [Parameter()]
        [object]$ShouldProcessTarget,

        [Parameter()]
        [switch]$StopOnError
    )

    $results = [System.Collections.Generic.List[pscustomobject]]::new()
    $catalogById = @{}
    foreach ($c in $Catalog) { $catalogById[$c.Id] = $c }

    foreach ($snap in $SnapshotControls) {
        $entry = $catalogById[$snap.id]

        $logRecord = [ordered]@{
            timestampUtc   = (Get-Date).ToUniversalTime().ToString('o')
            id             = $snap.id
            workload       = $snap.workload
            previousValue  = $null
            attemptedValue = $snap.currentValue
            result         = $null
            errorMessage   = $null
        }

        if (-not $entry) {
            $logRecord.result = $script:ResultStatus.Failed
            $logRecord.errorMessage = "No control named '$($snap.id)' exists in the currently loaded catalog; cannot restore it."
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $snap.id; Status = $script:ResultStatus.Failed; PreviousValue = $null; AppliedValue = $null; Message = $logRecord.errorMessage })
            if ($StopOnError) { throw "Stopping (StopOnError): $($snap.id) - $($logRecord.errorMessage)" }
            continue
        }

        if (-not $entry.Automatable) {
            $message = if ($entry.ManualInstructions) { $entry.ManualInstructions } else { 'No automated remediation is implemented for this control.' }
            $logRecord.result = $script:ResultStatus.SkippedManual
            $logRecord.errorMessage = $message
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = $script:ResultStatus.SkippedManual; PreviousValue = $null; AppliedValue = $null; Message = $message })
            continue
        }

        $snapHasError = $snap.PSObject.Properties['error'] -and $snap.error
        if ($snapHasError -or $null -eq $snap.currentValue) {
            $message = if ($snapHasError) {
                "This snapshot never captured a valid value for this control (the audit that produced it failed to read it: $($snap.error)); nothing to restore it to."
            }
            else {
                "This snapshot recorded a null/empty value for this control; nothing to restore it to."
            }
            $logRecord.result = 'Skipped-NoData'
            $logRecord.errorMessage = $message
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = 'Skipped-NoData'; PreviousValue = $null; AppliedValue = $null; Message = $message })
            continue
        }

        $target = "$($entry.Id) ($($entry.Workload))"
        $action = "Restore recorded value from snapshot: $($snap.currentValue | ConvertTo-Json -Compress -Depth 10)"
        $shouldProceed = $true
        if ($ShouldProcessTarget -and ($ShouldProcessTarget | Get-Member -Name ShouldProcess -ErrorAction SilentlyContinue)) {
            $shouldProceed = $ShouldProcessTarget.ShouldProcess($target, $action)
        }

        if (-not $shouldProceed -or $WhatIfMode) {
            $logRecord.result = 'Skipped-WhatIf'
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = 'Skipped-WhatIf'; PreviousValue = $null; AppliedValue = $null; Message = 'Skipped due to -WhatIf.' })
            continue
        }

        try {
            $setResult = & $entry.SetCommand -DesiredValue $snap.currentValue -CurrentValue $null
            $logRecord.result = $setResult.Status
            $logRecord.errorMessage = $setResult.Message
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add($setResult)
        }
        catch {
            $logRecord.result = $script:ResultStatus.Failed
            $logRecord.errorMessage = $_.Exception.Message
            Write-BaselineChangeLogEntry -Path $ChangeLogPath -Entry $logRecord
            $results.Add([pscustomobject]@{ Id = $entry.Id; Status = $script:ResultStatus.Failed; PreviousValue = $null; AppliedValue = $null; Message = $_.Exception.Message })
            if ($StopOnError) { throw "Stopping (StopOnError): $($entry.Id) - $($_.Exception.Message)" }
        }
    }

    return ,$results.ToArray()
}

# ---------------------------------------------------------------------------
# Backup / snapshot
# ---------------------------------------------------------------------------

function Save-BaselineSnapshot {
    <#
    .SYNOPSIS
        Writes a timestamped JSON backup/snapshot file from audit results.
    .PARAMETER AuditResults
        Results from Invoke-BaselineControlAudit.
    .PARAMETER Path
        File path to write to.
    .PARAMETER SourceMode
        'Audit', 'Apply-PreChange', or 'Apply-PostChange' - recorded for context.
    .PARAMETER ConfigSchemaVersion
        The baseline config's schemaVersion, recorded for traceability.
    .EXAMPLE
        Save-BaselineSnapshot -AuditResults $audit -Path ./backups/backup_...json -SourceMode Audit -ConfigSchemaVersion '1.0'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject[]]$AuditResults,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [ValidateSet('Audit', 'Apply-PreChange', 'Apply-PostChange')]
        [string]$SourceMode,

        [Parameter(Mandatory)]
        [string]$ConfigSchemaVersion
    )

    $snapshot = [pscustomobject]@{
        snapshotSchemaVersion    = $script:SnapshotSchemaVersion
        capturedAtUtc            = (Get-Date).ToUniversalTime().ToString('o')
        mode                     = $SourceMode
        baselineConfigSchemaVersion = $ConfigSchemaVersion
        controls                 = @($AuditResults | ForEach-Object {
            [pscustomobject]@{
                id             = $_.Id
                workload       = $_.Workload
                description    = $_.Description
                automatable    = $_.Automatable
                currentValue   = $_.CurrentValue
                desiredValue   = $_.DesiredValue
                compliant      = $_.Compliant
                error          = $_.Error
            }
        })
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force -WhatIf:$false | Out-Null }

    $snapshot | ConvertTo-Json -Depth 25 | Set-Content -LiteralPath $Path -Encoding utf8 -WhatIf:$false
    return $Path
}

function Import-BaselineSnapshot {
    <#
    .SYNOPSIS
        Loads and validates a snapshot/backup file for use by Restore mode.
    .DESCRIPTION
        Refuses to load a snapshot whose snapshotSchemaVersion doesn't match what
        this build of the toolkit understands, rather than silently applying a
        partially-understood file.
    .PARAMETER Path
        Path to the snapshot JSON file.
    .EXAMPLE
        Import-BaselineSnapshot -Path ./backups/backup_2026-09-17T14-30-00Z.json
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Backup/snapshot file not found: $Path"
    }

    try {
        $snapshot = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -Depth 25 -ErrorAction Stop
    }
    catch {
        throw "Backup/snapshot file '$Path' is not valid JSON: $($_.Exception.Message)"
    }

    if (-not $snapshot.PSObject.Properties['snapshotSchemaVersion']) {
        throw "Backup/snapshot file '$Path' has no 'snapshotSchemaVersion' field; refusing to restore from a file this toolkit cannot confirm the shape of."
    }
    if ($snapshot.snapshotSchemaVersion -ne $script:SnapshotSchemaVersion) {
        throw "Backup/snapshot file '$Path' has snapshotSchemaVersion '$($snapshot.snapshotSchemaVersion)', but this build of the toolkit expects '$script:SnapshotSchemaVersion'. Refusing to restore from a file whose schema may not be fully understood."
    }
    if (-not $snapshot.PSObject.Properties['controls'] -or -not $snapshot.controls) {
        throw "Backup/snapshot file '$Path' has no 'controls' array; nothing to restore."
    }

    return $snapshot
}

# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

function Format-BaselineValueForDisplay {
    <#
    .SYNOPSIS
        Renders any value (scalar, array, object, $null) as a short inline string
        suitable for a Markdown table cell.
    .PARAMETER Value
        Value to render.
    .EXAMPLE
        Format-BaselineValueForDisplay -Value @{min=2;max=4}
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()]
        [AllowNull()]
        [object]$Value
    )
    if ($null -eq $Value) { return '_(none)_' }
    if ($Value -is [bool]) { return $Value.ToString() }
    if ($Value -is [string]) { return $Value }
    $json = $Value | ConvertTo-Json -Compress -Depth 10
    # Escape pipe characters so the value doesn't break the Markdown table.
    return ($json -replace '\|', '\|')
}

function Export-BaselineMarkdownReport {
    <#
    .SYNOPSIS
        Writes a Markdown compliance report, one row per control.
    .PARAMETER AuditResults
        Results from Invoke-BaselineControlAudit.
    .PARAMETER Path
        File path to write to.
    .PARAMETER Title
        Report title/heading.
    .PARAMETER ApplyResults
        Optional. If supplied (Apply mode's post-change report), adds Action/Result columns.
    .EXAMPLE
        Export-BaselineMarkdownReport -AuditResults $audit -Path ./reports/pre-change_....md -Title 'Pre-Change Audit'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject[]]$AuditResults,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter()]
        [pscustomobject[]]$ApplyResults
    )

    $applyById = @{}
    if ($ApplyResults) { foreach ($r in $ApplyResults) { $applyById[$r.Id] = $r } }

    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add("# $Title")
    $lines.Add('')
    $lines.Add("Generated: $((Get-Date).ToUniversalTime().ToString('o'))")
    $lines.Add('')
    $errorCount = @($AuditResults | Where-Object { $_.Error }).Count
    $nonCompliantCount = @($AuditResults | Where-Object { $_.Compliant -eq $false }).Count
    $lines.Add("Controls evaluated: $($AuditResults.Count) | Non-compliant: $nonCompliantCount | Read errors: $errorCount")
    $lines.Add('')

    if ($ApplyResults) {
        $lines.Add('| Id | Workload | Setting | Current | Desired | Compliant | Automatable | Action Taken | Result |')
        $lines.Add('|---|---|---|---|---|---|---|---|---|')
    }
    else {
        $lines.Add('| Id | Workload | Setting | Current | Desired | Compliant | Automatable | Notes |')
        $lines.Add('|---|---|---|---|---|---|---|---|')
    }

    foreach ($r in $AuditResults) {
        $compliantText = if ($null -eq $r.Compliant) { 'Unknown' } elseif ($r.Compliant) { 'Yes' } else { 'No' }
        $current = if ($r.Error) { "_error: $($r.Error)_" } else { Format-BaselineValueForDisplay -Value $r.CurrentValue }
        $desired = Format-BaselineValueForDisplay -Value $r.DesiredValue

        if ($ApplyResults) {
            $applyResult = $applyById[$r.Id]
            $action = if ($applyResult) { $applyResult.Status } else { 'N/A' }
            $resultMsg = if ($applyResult -and $applyResult.Message) { $applyResult.Message } else { '' }
            $lines.Add("| $($r.Id) | $($r.Workload) | $($r.Description) | $current | $desired | $compliantText | $($r.Automatable) | $action | $resultMsg |")
        }
        else {
            $notes = if (-not $r.Automatable) { "Manual: $($r.ManualInstructions)" } else { '' }
            $lines.Add("| $($r.Id) | $($r.Workload) | $($r.Description) | $current | $desired | $compliantText | $($r.Automatable) | $notes |")
        }
    }

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force -WhatIf:$false | Out-Null }

    ($lines -join "`n") | Set-Content -LiteralPath $Path -Encoding utf8 -WhatIf:$false
    return $Path
}

function Export-BaselineHtmlReport {
    <#
    .SYNOPSIS
        Writes a minimal, dependency-free HTML compliance report alongside the
        Markdown report.
    .PARAMETER AuditResults
        Results from Invoke-BaselineControlAudit.
    .PARAMETER Path
        File path to write to.
    .PARAMETER Title
        Report title.
    .PARAMETER ApplyResults
        Optional post-change results (see Export-BaselineMarkdownReport).
    .EXAMPLE
        Export-BaselineHtmlReport -AuditResults $audit -Path ./reports/pre-change_....html -Title 'Pre-Change Audit'
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [pscustomobject[]]$AuditResults,

        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$Title,

        [Parameter()]
        [pscustomobject[]]$ApplyResults
    )

    $applyById = @{}
    if ($ApplyResults) { foreach ($r in $ApplyResults) { $applyById[$r.Id] = $r } }

    $rowsHtml = foreach ($r in $AuditResults) {
        $compliantText = if ($null -eq $r.Compliant) { 'Unknown' } elseif ($r.Compliant) { 'Yes' } else { 'No' }
        $current = if ($r.Error) { "error: $($r.Error)" } else { Format-BaselineValueForDisplay -Value $r.CurrentValue }
        $desired = Format-BaselineValueForDisplay -Value $r.DesiredValue
        $extra = if ($ApplyResults) {
            $applyResult = $applyById[$r.Id]
            $action = if ($applyResult) { $applyResult.Status } else { 'N/A' }
            $msg = if ($applyResult) { $applyResult.Message } else { '' }
            "<td>$([System.Net.WebUtility]::HtmlEncode($action))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$msg))</td>"
        }
        else {
            $notes = if (-not $r.Automatable) { "Manual: $($r.ManualInstructions)" } else { '' }
            "<td>$([System.Net.WebUtility]::HtmlEncode($notes))</td>"
        }
        "<tr><td>$([System.Net.WebUtility]::HtmlEncode($r.Id))</td><td>$([System.Net.WebUtility]::HtmlEncode($r.Workload))</td><td>$([System.Net.WebUtility]::HtmlEncode($r.Description))</td><td>$([System.Net.WebUtility]::HtmlEncode($current))</td><td>$([System.Net.WebUtility]::HtmlEncode($desired))</td><td>$compliantText</td><td>$($r.Automatable)</td>$extra</tr>"
    }

    $extraHeader = if ($ApplyResults) { '<th>Action Taken</th><th>Result</th>' } else { '<th>Notes</th>' }

    $html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>$([System.Net.WebUtility]::HtmlEncode($Title))</title>
<style>
body { font-family: -apple-system, Segoe UI, Arial, sans-serif; margin: 2rem; color: #1a1a1a; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ccc; padding: 6px 10px; text-align: left; font-size: 0.9rem; vertical-align: top; }
th { background: #f2f2f2; }
tr:nth-child(even) { background: #fafafa; }
</style>
</head>
<body>
<h1>$([System.Net.WebUtility]::HtmlEncode($Title))</h1>
<p>Generated: $((Get-Date).ToUniversalTime().ToString('o'))</p>
<table>
<thead><tr><th>Id</th><th>Workload</th><th>Setting</th><th>Current</th><th>Desired</th><th>Compliant</th><th>Automatable</th>$extraHeader</tr></thead>
<tbody>
$($rowsHtml -join "`n")
</tbody>
</table>
</body>
</html>
"@

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force -WhatIf:$false | Out-Null }

    $html | Set-Content -LiteralPath $Path -Encoding utf8 -WhatIf:$false
    return $Path
}

function Write-BaselineChangeLogEntry {
    <#
    .SYNOPSIS
        Appends one JSON Lines record to the structured change log.
    .PARAMETER Path
        Change log file path.
    .PARAMETER Entry
        Ordered hashtable/object with the record fields.
    .EXAMPLE
        Write-BaselineChangeLogEntry -Path ./reports/changelog_....jsonl -Entry $record
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [object]$Entry
    )
    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force -WhatIf:$false | Out-Null }

    ($Entry | ConvertTo-Json -Compress -Depth 10) | Add-Content -LiteralPath $Path -Encoding utf8 -WhatIf:$false
}

function Get-BaselineTimestampedPath {
    <#
    .SYNOPSIS
        Builds a collision-free, UTC-timestamped output file path and ensures its
        parent directory exists.
    .PARAMETER Directory
        Target directory.
    .PARAMETER Prefix
        File name prefix (e.g. 'backup', 'pre-change').
    .PARAMETER Extension
        File extension without a leading dot (e.g. 'json', 'md').
    .EXAMPLE
        Get-BaselineTimestampedPath -Directory ./backups -Prefix backup -Extension json
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory)]
        [string]$Directory,

        [Parameter(Mandatory)]
        [string]$Prefix,

        [Parameter(Mandatory)]
        [string]$Extension
    )
    if (-not (Test-Path -LiteralPath $Directory)) { New-Item -ItemType Directory -Path $Directory -Force -WhatIf:$false | Out-Null }
    $stamp = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH-mm-ssZ')
    return (Join-Path $Directory "${Prefix}_${stamp}.${Extension}")
}

Export-ModuleMember -Function @(
    'Import-BaselineConfig'
    'Test-BaselineConfigSemantics'
    'Get-BaselineControlConnection'
    'Get-BaselineControlExtraConnections'
    'Get-BaselineControlCatalog'
    'Get-BaselineSubscribedSkuCache'
    'Test-TenantServicePlan'
    'Compare-BaselineValueDeep'
    'Get-BaselinePropertyMap'
    'Test-BaselineCompliance'
    'Test-BaselineRequiredModule'
    'Install-BaselineRequiredModule'
    'Assert-BaselineRequiredModules'
    'Get-BaselineConnectionOrder'
    'Test-BaselineWorkloadConnected'
    'Set-BaselineWorkloadConnectedState'
    'Connect-BaselineWorkload'
    'Disconnect-BaselineWorkload'
    'Invoke-BaselineControlAudit'
    'Test-BaselineApplyReadiness'
    'Invoke-BaselineControlApply'
    'Invoke-BaselineControlRestore'
    'Save-BaselineSnapshot'
    'Import-BaselineSnapshot'
    'Format-BaselineValueForDisplay'
    'Export-BaselineMarkdownReport'
    'Export-BaselineHtmlReport'
    'Write-BaselineChangeLogEntry'
    'Get-BaselineTimestampedPath'
)
