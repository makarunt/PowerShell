#Requires -Version 7.0
<#
    ConditionalAccessControls.psm1

    Get-/Set- function pairs for the toolkit's Conditional Access (CA) controls.
    Requires Microsoft.Graph (v2+) to be connected before use (Connect-BaselineWorkload
    -Connection Graph) with at least Policy.ReadWrite.ConditionalAccess,
    Policy.Read.All (also covers reading whether Security Defaults is enabled,
    via Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy - see
    Test-BaselineCASecurityDefaultsEnabled), Group.ReadWrite.All,
    Organization.Read.All, and Application.Read.All.

    NON-NEGOTIABLE: every policy this module creates or updates is set to
    state = 'enabledForReportingButNotEnforced' ("report-only"). No code path in
    this module ever sets a CA policy's state to 'enabled'. Enabling a policy this
    toolkit created is a manual, deliberate step the tenant owner takes later, in
    the Entra admin center, after reviewing the report-only sign-in logs - see
    README.md.

    Licensing: this whole module is gated on Entra ID P1 (Tier 1 controls) or P2
    (Tier 2 controls), via the shared Test-TenantServicePlan helper in
    BaselineCore.psm1. The other five control modules (EntraId/ExchangeOnline/
    Teams/SharePointOnline directory-settings controls) run on Entra ID Free and
    must never be made to depend on anything gated in here.

    Idempotency: every policy this module manages is named with the fixed prefix
    "[M365 Baseline] " - Get- matches by exact display name. Before *creating* a
    toolkit-owned policy (never before updating one that already exists), this
    module scans every OTHER existing CA policy for a heuristic match and skips
    creation on a hit (Skipped-PotentialOverlap) unless forceCreateDespiteOverlap
    is set for that control in config.
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

$script:CAPolicyNamePrefix = '[M365 Baseline] '
$script:CAEmergencyGroupDisplayName = 'M365 Baseline - Emergency Access Accounts (DO NOT DELETE)'

# The only state this module is ever allowed to write. Do not add a way to
# override this, even behind a flag - see the module header.
$script:CAReportOnlyState = 'enabledForReportingButNotEnforced'

$script:CATier1ServicePlans = @('AAD_PREMIUM')
$script:CATier2ServicePlans = @('AAD_PREMIUM_P2')

$script:CAAdminRoleDisplayNames = @(
    'Global Administrator', 'Privileged Role Administrator', 'Application Administrator',
    'Cloud Application Administrator', 'Authentication Administrator', 'Privileged Authentication Administrator',
    'Security Administrator', 'Exchange Administrator', 'SharePoint Administrator', 'User Administrator',
    'Helpdesk Administrator', 'Conditional Access Administrator', 'Billing Administrator', 'Password Administrator'
)

# Microsoft's own well-known, stable AppId for the first-party "Microsoft Azure
# Management" application - the same across every tenant. Still verified against
# Get-MgServicePrincipal before use (see Test-BaselineCAAzureManagementAppResolvable)
# rather than trusted blindly, in case it's ever absent in a given tenant.
$script:CAAzureManagementAppId = '797f4846-ba00-4fd7-ba43-dac1f8f63013'

# ---------------------------------------------------------------------------
# Per-run caches (module-scoped; reset whenever this module is re-imported with
# -Force, i.e. once per Invoke-M365Baseline.ps1 run - see BaselineCore.psm1's
# Get-BaselineSubscribedSkuCache for the same pattern/rationale)
# ---------------------------------------------------------------------------

$script:CAPolicyListCache = $null
$script:CAEmergencyGroupIdCache = $null
$script:CAAdminRoleIdCache = $null
$script:CASecurityDefaultsEnabledCache = $null

# ---------------------------------------------------------------------------
# Shared engine: licensing, policy lookup, overlap detection, emergency group
# ---------------------------------------------------------------------------

function Test-BaselineCATierAvailable {
    <#
    .SYNOPSIS
        Checks whether the tenant's licensing satisfies a given CA control tier.
    .DESCRIPTION
        Tier 1 needs Entra ID P1 - satisfied by AAD_PREMIUM OR AAD_PREMIUM_P2 (a
        P2 tenant also gets everything Tier 1 offers). Tier 2 needs Entra ID P2
        specifically. Built on the shared Test-TenantServicePlan helper (see
        BaselineCore.psm1) so Get-MgSubscribedSku is still only called once per
        run no matter how many controls check their tier.
    .PARAMETER Tier
        1 or 2.
    .EXAMPLE
        Test-BaselineCATierAvailable -Tier 1
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [ValidateSet(1, 2)]
        [int]$Tier
    )
    if ($Tier -eq 1) {
        return (Test-TenantServicePlan -ServicePlanNames $script:CATier1ServicePlans) -or (Test-TenantServicePlan -ServicePlanNames $script:CATier2ServicePlans)
    }
    return Test-TenantServicePlan -ServicePlanNames $script:CATier2ServicePlans
}

function Test-BaselineCASecurityDefaultsEnabled {
    <#
    .SYNOPSIS
        Checks whether this tenant currently has Microsoft Entra Security
        Defaults enabled.
    .DESCRIPTION
        Confirmed against current Microsoft documentation: creating a
        Conditional Access policy - even report-only, exactly what this
        module ever does - permanently removes the ability to re-enable
        Security Defaults afterward. The "Manage security defaults" toggle
        stays unavailable while ANY Conditional Access policy exists in the
        tenant, in ANY state, until every one of them (this module's own
        included) is deleted. That's a one-way door for the tenant, distinct
        from and in addition to this module's report-only-never-enforced
        guarantee, so Set- refuses to CREATE a new toolkit-owned policy
        while Security Defaults remains on rather than silently taking away
        that option. Cached per run (reset on the module's next -Force
        import) since every CA control's Get-/Set- can hit this on the
        "policy doesn't exist yet" path.
    .EXAMPLE
        Test-BaselineCASecurityDefaultsEnabled
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    if ($null -eq $script:CASecurityDefaultsEnabledCache) {
        $policy = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy -ErrorAction Stop
        $script:CASecurityDefaultsEnabledCache = [bool]$policy.IsEnabled
    }
    return $script:CASecurityDefaultsEnabledCache
}

function Get-BaselineCAPolicyListCache {
    <#
    .SYNOPSIS
        Internal: returns this run's cached list of every existing CA policy
        (toolkit-owned or not), fetching it once and reusing it.
    .PARAMETER Refresh
        Forces a fresh Get-MgIdentityConditionalAccessPolicy call. Set-
        functions pass this after a create/update so a later read in the same
        run (e.g. Apply's post-change audit) sees the change.
    .EXAMPLE
        Get-BaselineCAPolicyListCache
    #>
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter()]
        [switch]$Refresh
    )
    if ($Refresh -or $null -eq $script:CAPolicyListCache) {
        $script:CAPolicyListCache = @(Get-MgIdentityConditionalAccessPolicy -All -ErrorAction Stop)
    }
    # Plain return, deliberately NOT the usual ",$array" empty/single-element-safe
    # idiom used elsewhere in this toolkit (e.g. Get-BaselineSubscribedSkuCache):
    # every caller here pipes this call's output directly (Find-BaselineCAPolicyByName)
    # or foreach-es over the call expression without assigning it to a variable
    # first (Find-BaselineCAOverlap) - in both of those call shapes, ",$array"
    # deposits the WHOLE array as a single pipeline/loop item instead of
    # enumerating its elements, which is the opposite of what's needed here.
    # The idiom is only safe for "$x = Get-Foo" assignment call sites, which
    # this function has none of.
    return $script:CAPolicyListCache
}

function Find-BaselineCAPolicyByName {
    <#
    .SYNOPSIS
        Exact-display-name lookup for a toolkit-owned policy - the whole basis
        of this module's idempotency.
    .PARAMETER DisplayName
        Exact policy display name, e.g. '[M365 Baseline] Require MFA for all users'.
    .EXAMPLE
        Find-BaselineCAPolicyByName -DisplayName '[M365 Baseline] Require MFA for all users'
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory)]
        [string]$DisplayName
    )
    return (Get-BaselineCAPolicyListCache | Where-Object { [string]$_.DisplayName -eq $DisplayName } | Select-Object -First 1)
}

function Find-BaselineCAOverlap {
    <#
    .SYNOPSIS
        Scans every CA policy this toolkit does NOT own for a heuristic match,
        used to avoid creating a duplicate/conflicting policy.
    .DESCRIPTION
        Only ever called before CREATING a toolkit-owned policy (never before
        updating one that already exists - an existing toolkit-owned policy is
        by definition not an "overlap", it's the same policy being converged).
    .PARAMETER Predicate
        Scriptblock taking one policy object, returning $true on a heuristic match.
    .EXAMPLE
        Find-BaselineCAOverlap -Predicate { param($p) $p.GrantControls.BuiltInControls -contains 'mfa' }
    #>
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory)]
        [scriptblock]$Predicate
    )
    foreach ($p in (Get-BaselineCAPolicyListCache)) {
        if ([string]$p.DisplayName -like "$($script:CAPolicyNamePrefix)*") { continue }
        if (& $Predicate $p) { return $p }
    }
    return $null
}

function Get-BaselineCAEmergencyAccessGroupId {
    <#
    .SYNOPSIS
        Resolves the id of the placeholder emergency-access group, optionally
        creating it (empty) if it doesn't exist yet.
    .DESCRIPTION
        Read-only by default (-CreateIfMissing not set) so that Audit/Get- never
        creates anything - only Set- (Apply) passes -CreateIfMissing. When the
        group doesn't exist and creation wasn't requested, returns $null; every
        caller treats a $null group id as "the emergency-access exclusion cannot
        be verified/applied yet", which correctly reports as non-compliant
        rather than silently skipping the check.
    .PARAMETER CreateIfMissing
        Create the group (as an empty security group) if no group with the
        exact expected display name exists yet.
    .PARAMETER Refresh
        Forces a fresh lookup even if a cached id exists.
    .EXAMPLE
        Get-BaselineCAEmergencyAccessGroupId -CreateIfMissing
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter()]
        [switch]$CreateIfMissing,

        [Parameter()]
        [switch]$Refresh
    )
    if (-not $Refresh -and $script:CAEmergencyGroupIdCache) {
        return $script:CAEmergencyGroupIdCache
    }
    $escapedName = $script:CAEmergencyGroupDisplayName.Replace("'", "''")
    $group = Get-MgGroup -Filter "displayName eq '$escapedName'" -ConsistencyLevel eventual -CountVariable groupCount -All -ErrorAction Stop | Select-Object -First 1
    if (-not $group -and $CreateIfMissing) {
        $group = New-MgGroup -DisplayName $script:CAEmergencyGroupDisplayName -MailEnabled:$false -MailNickname 'M365BaselineEmergencyAccess' -SecurityEnabled:$true -ErrorAction Stop
    }
    if ($group) {
        $script:CAEmergencyGroupIdCache = [string]$group.Id
    }
    return $script:CAEmergencyGroupIdCache
}

function Get-BaselineCAAdminRoleTemplateIds {
    <#
    .SYNOPSIS
        Resolves CA-RequireMfaAdminRoles' 14 directory role template ids
        dynamically at runtime, never from hardcoded GUIDs.
    .DESCRIPTION
        Throws a clear, actionable error naming exactly which role(s) could not
        be resolved rather than silently building a policy that targets fewer
        roles than intended.
    .EXAMPLE
        Get-BaselineCAAdminRoleTemplateIds
    #>
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter()]
        [switch]$Refresh
    )
    # Plain return, not the usual ",$array" empty/single-element-array-safe idiom:
    # every caller of this function either pipes its output directly or
    # foreach-es over the call expression without an intermediate assignment
    # (see Get-BaselineCAPolicyListCache's comment below for why that idiom
    # actively breaks those specific call shapes rather than protecting them).
    if (-not $Refresh -and $script:CAAdminRoleIdCache) {
        return $script:CAAdminRoleIdCache
    }
    # The "$_ -and" guard defends against a stray null entry in the pipeline
    # (harmless either way). The real bug this whole block works around:
    # Set-StrictMode -Version Latest throws "The property 'DisplayName' cannot
    # be found on this object" on $templates.DisplayName member-enumeration
    # when $templates is a genuinely EMPTY array - confirmed directly, this is
    # not limited to $null elements. That's exactly the shape a tenant missing
    # one of these roles (or a test mocking Get-MgDirectoryRoleTemplate to
    # return none) produces, which without this fix crashes before ever
    # reaching this function's own, more useful "could not resolve" error
    # below. ForEach-Object -Property, unlike dotted member-enumeration, never
    # touches .DisplayName at all when there are zero elements to iterate.
    $templates = @(Get-MgDirectoryRoleTemplate -All -ErrorAction Stop | Where-Object { $_ -and ($script:CAAdminRoleDisplayNames -contains [string]$_.DisplayName) })
    $resolvedNames = @($templates | ForEach-Object { [string]$_.DisplayName })
    $missing = @($script:CAAdminRoleDisplayNames | Where-Object { $resolvedNames -notcontains $_ })
    if ($missing.Count -gt 0) {
        throw "CA-RequireMfaAdminRoles: could not resolve directory role template(s) via Get-MgDirectoryRoleTemplate: $($missing -join ', '). These are expected to be standard Entra ID built-in roles - verify they exist in this tenant."
    }
    $script:CAAdminRoleIdCache = [string[]]$templates.Id
    return $script:CAAdminRoleIdCache
}

function Test-BaselineCAAzureManagementAppResolvable {
    <#
    .SYNOPSIS
        Confirms the well-known "Microsoft Azure Management" first-party app
        actually resolves in this tenant before a policy is built to reference it.
    .EXAMPLE
        Test-BaselineCAAzureManagementAppResolvable
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    $sp = Get-MgServicePrincipal -Filter "appId eq '$script:CAAzureManagementAppId'" -ErrorAction Stop
    return [bool]$sp
}

function Test-BaselineCAGrantControlsMatch {
    <#
    .SYNOPSIS
        Order-insensitive comparison of a live policy's GrantControls against
        an expected set of built-in controls and operator.
    .PARAMETER Policy
        The live CA policy object (from Get-MgIdentityConditionalAccessPolicy).
    .PARAMETER ExpectedControls
        Expected BuiltInControls values, e.g. @('mfa').
    .PARAMETER ExpectedOperator
        'OR' or 'AND'. Defaults to 'OR'.
    .EXAMPLE
        Test-BaselineCAGrantControlsMatch -Policy $policy -ExpectedControls @('mfa')
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [object]$Policy,

        [Parameter(Mandatory)]
        [string[]]$ExpectedControls,

        [Parameter()]
        [string]$ExpectedOperator = 'OR'
    )
    $actual = @($Policy.GrantControls.BuiltInControls | ForEach-Object { [string]$_ })
    $expectedSorted = @($ExpectedControls | Sort-Object)
    $actualSorted = @($actual | Sort-Object)
    if (@(Compare-Object -ReferenceObject $expectedSorted -DifferenceObject $actualSorted)) { return $false }
    return [string]$Policy.GrantControls.Operator -ieq $ExpectedOperator
}

function Test-BaselineCAEmergencyGroupExcluded {
    <#
    .SYNOPSIS
        Checks whether the emergency-access group is present in a live policy's
        Conditions.Users.ExcludeGroups.
    .PARAMETER Policy
        The live CA policy object.
    .PARAMETER EmergencyGroupId
        Id of the emergency-access group, or $null if it doesn't exist yet (in
        which case this always returns $false - see Get-BaselineCAEmergencyAccessGroupId).
    .EXAMPLE
        Test-BaselineCAEmergencyGroupExcluded -Policy $policy -EmergencyGroupId $groupId
    #>
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory)]
        [object]$Policy,

        [Parameter()]
        [AllowNull()]
        [string]$EmergencyGroupId
    )
    if ([string]::IsNullOrWhiteSpace($EmergencyGroupId)) { return $false }
    $excluded = @($Policy.Conditions.Users.ExcludeGroups | ForEach-Object { [string]$_ })
    return $excluded -contains $EmergencyGroupId
}

function Get-BaselineCAControlState {
    <#
    .SYNOPSIS
        Generic Get- engine shared by every CA control's public Get-<Id>State
        wrapper.
    .DESCRIPTION
        Order of checks: (1) tier/license gate - a gated-out control reports
        Value = $null ("Unknown", the same convention EntraID's audit-only
        controls use), never an error; (2) does a policy with this exact display
        name exist; (3) does it match state=report-only, the emergency-group
        exclusion, and the control's own condition/grant shape.

        When the toolkit-owned policy doesn't exist yet, this also runs the same
        (read-only) overlap scan Set- uses before deciding whether to create one,
        so Audit can tell you up front whether Apply would actually create
        anything here or just skip it as redundant with an existing policy -
        without this, "Compliant: No" during Audit couldn't distinguish "genuinely
        missing" from "would be skipped as redundant," which only became
        knowable once Set- actually ran. Purely informational: Find-BaselineCAOverlap
        only scans, it never creates or changes anything, so this is safe to run
        during Audit. Doesn't know about this control's forceCreateDespiteOverlap
        config setting (Get- isn't passed that - only Set- is), so the note is
        phrased to stay accurate whichever way that's set.
    .PARAMETER Spec
        Hashtable: Id, DisplayName, Tier, ComplianceCheck (scriptblock($policy) -> bool,
        checking only conditions/grantControls - state and emergency-group
        exclusion are checked here, once, for every control), OverlapPredicate
        (scriptblock($policy) -> bool, reused here purely for the Audit-time note).
    .EXAMPLE
        Get-BaselineCAControlState -Spec $spec
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Spec
    )
    if (-not (Test-BaselineCATierAvailable -Tier $Spec.Tier)) {
        return [pscustomobject]@{ Id = $Spec.Id; Value = $null; Detail = "Tenant does not have the Entra ID license Tier $($Spec.Tier) requires." }
    }
    $policy = Find-BaselineCAPolicyByName -DisplayName $Spec.DisplayName
    if (-not $policy) {
        $overlap = Find-BaselineCAOverlap -Predicate $Spec.OverlapPredicate
        # Checked here (not just in Set-) so an admin sees this before ever
        # running Apply, not as a surprise afterward - see
        # Test-BaselineCASecurityDefaultsEnabled for why it matters.
        $securityDefaultsNote = if (Test-BaselineCASecurityDefaultsEnabled) {
            " This tenant has Security Defaults enabled - Apply will NOT create this report-only policy while that remains on, since creating any Conditional Access policy (even report-only) permanently blocks re-enabling Security Defaults until every Conditional Access policy in the tenant, including this one, is deleted. Disable Security Defaults first (Entra admin center > Identity > Overview > Properties > Manage security defaults) if you want this control automated."
        }
        else { '' }
        $detail = if ($overlap) {
            "Policy '$($Spec.DisplayName)' does not exist yet. An existing, non-toolkit-owned policy ('$($overlap.DisplayName)', id $($overlap.Id)) heuristically overlaps with this control's intent - Apply will skip creating this one (Skipped-PotentialOverlap) unless forceCreateDespiteOverlap is set for it in config.$securityDefaultsNote"
        }
        else {
            "Policy '$($Spec.DisplayName)' does not exist yet.$securityDefaultsNote"
        }
        return [pscustomobject]@{ Id = $Spec.Id; Value = $false; Detail = $detail }
    }
    $emergencyGroupId = Get-BaselineCAEmergencyAccessGroupId
    $compliant = ([string]$policy.State -eq $script:CAReportOnlyState) -and
        (Test-BaselineCAEmergencyGroupExcluded -Policy $policy -EmergencyGroupId $emergencyGroupId) -and
        [bool](& $Spec.ComplianceCheck $policy)
    return [pscustomobject]@{ Id = $Spec.Id; Value = [bool]$compliant; Detail = "Existing policy '$($Spec.DisplayName)' (id $($policy.Id)) found." }
}

function Set-BaselineCAControlState {
    <#
    .SYNOPSIS
        Generic Set- engine shared by every CA control's public Set-<Id>State
        wrapper.
    .DESCRIPTION
        Order of checks: (1) already compliant per the passed-in CurrentValue ->
        no-op; (2) tier/license gate -> Skipped-LicenseInsufficient; (3) ensure
        the emergency-access group exists (creating it if needed - this is the
        only code path in this module allowed to create it); (4) if no
        toolkit-owned policy exists yet, scan for overlap with a non-owned
        policy and skip creation on a hit unless ForceCreateDespiteOverlap;
        (5) create or update, always with state = report-only and the
        emergency-access group excluded.
    .PARAMETER Spec
        Hashtable: Id, DisplayName, Tier, BuildConditions (scriptblock($emergencyGroupId)
        -> conditions body hashtable), GrantControls (hashtable), OverlapPredicate
        (scriptblock($policy) -> bool), ComplianceCheck (scriptblock($policy) -> bool).
    .PARAMETER CurrentValue
        Pre-fetched compliance bool from Get-BaselineCAControlState, or $null.
    .PARAMETER ForceCreateDespiteOverlap
        Skips the overlap check for this control's creation. Driven by the
        control's forceCreateDespiteOverlap config field.
    .EXAMPLE
        Set-BaselineCAControlState -Spec $spec -CurrentValue $false
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [hashtable]$Spec,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )

    if ($null -ne $CurrentValue -and [bool]$CurrentValue -eq $true) {
        return [pscustomobject]@{ Id = $Spec.Id; Status = 'Success'; PreviousValue = $true; AppliedValue = $true; Message = "Already compliant (no-op). Report-only: state=$($script:CAReportOnlyState), never enforced by this toolkit." }
    }

    if (-not (Test-BaselineCATierAvailable -Tier $Spec.Tier)) {
        $planName = if ($Spec.Tier -eq 1) { 'AAD_PREMIUM (Entra ID P1) or AAD_PREMIUM_P2 (Entra ID P2)' } else { 'AAD_PREMIUM_P2 (Entra ID P2)' }
        return [pscustomobject]@{ Id = $Spec.Id; Status = 'Skipped-LicenseInsufficient'; PreviousValue = $CurrentValue; AppliedValue = $null; Message = "Tenant is missing the required service plan(s): $planName. This Conditional Access control was not created/updated." }
    }

    $emergencyGroupId = Get-BaselineCAEmergencyAccessGroupId -CreateIfMissing
    $policy = Find-BaselineCAPolicyByName -DisplayName $Spec.DisplayName

    if (-not $policy) {
        # Checked before the overlap scan, and before ever building/sending a
        # create request: creating ANY Conditional Access policy - even
        # report-only, all this module ever does - permanently blocks the
        # tenant from re-enabling Security Defaults later, until every CA
        # policy (this one included) is deleted again. That's a one-way door
        # this module must never open on its own, so there is deliberately no
        # override for it (unlike ForceCreateDespiteOverlap above) - see
        # Test-BaselineCASecurityDefaultsEnabled.
        if (Test-BaselineCASecurityDefaultsEnabled) {
            $message = 'Skipped - this tenant has Microsoft Entra Security Defaults enabled. Creating any Conditional Access policy, even report-only, permanently blocks re-enabling Security Defaults until every Conditional Access policy in the tenant is deleted, so this toolkit will not create one while Security Defaults remains on. Disable Security Defaults first (Entra admin center > Identity > Overview > Properties > Manage security defaults), then re-run Apply, if you want this control automated.'
            return [pscustomobject]@{ Id = $Spec.Id; Status = 'Skipped-SecurityDefaultsEnabled'; PreviousValue = $CurrentValue; AppliedValue = $null; Message = $message }
        }
        $overlap = Find-BaselineCAOverlap -Predicate $Spec.OverlapPredicate
        if ($overlap -and -not $ForceCreateDespiteOverlap) {
            $message = "Skipped - an existing, non-toolkit-owned policy ('$($overlap.DisplayName)', id $($overlap.Id)) already looks like it covers this. Review it manually; set forceCreateDespiteOverlap: true for this control in config/baseline.config.json if you still want this toolkit-owned report-only policy created alongside it."
            return [pscustomobject]@{ Id = $Spec.Id; Status = 'Skipped-PotentialOverlap'; PreviousValue = $CurrentValue; AppliedValue = $null; Message = $message }
        }
    }

    $conditions = & $Spec.BuildConditions $emergencyGroupId
    $body = @{
        displayName   = $Spec.DisplayName
        state         = $script:CAReportOnlyState
        conditions    = $conditions
        grantControls = $Spec.GrantControls
    }

    if (-not $policy) {
        New-MgIdentityConditionalAccessPolicy -BodyParameter $body -ErrorAction Stop | Out-Null
        Get-BaselineCAPolicyListCache -Refresh | Out-Null
        $message = "Created report-only policy '$($Spec.DisplayName)' - state=$($script:CAReportOnlyState), NOT enforced. Emergency-access group excluded."
        return [pscustomobject]@{ Id = $Spec.Id; Status = 'Created'; PreviousValue = $false; AppliedValue = $true; Message = $message }
    }

    Update-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId $policy.Id -BodyParameter $body -ErrorAction Stop | Out-Null
    Get-BaselineCAPolicyListCache -Refresh | Out-Null
    $message = "Updated existing policy '$($Spec.DisplayName)' to match the baseline definition - state=$($script:CAReportOnlyState), NOT enforced. Emergency-access group excluded."
    return [pscustomobject]@{ Id = $Spec.Id; Status = 'Updated'; PreviousValue = $false; AppliedValue = $true; Message = $message }
}

function Get-BaselineConditionalAccessSummary {
    <#
    .SYNOPSIS
        One-shot licensing/report-only summary for the orchestrator to print
        prominently whenever any ConditionalAccess-workload control is enabled.
    .DESCRIPTION
        Read-only: does not create the emergency-access group (passes no
        -CreateIfMissing), so calling this from Audit mode changes nothing.
    .EXAMPLE
        Get-BaselineConditionalAccessSummary
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    return [pscustomobject]@{
        Tier1Available   = Test-BaselineCATierAvailable -Tier 1
        Tier2Available   = Test-BaselineCATierAvailable -Tier 2
        ReportOnlyState  = $script:CAReportOnlyState
        EmergencyGroupId = Get-BaselineCAEmergencyAccessGroupId
        EmergencyGroupDisplayName = $script:CAEmergencyGroupDisplayName
    }
}

# ---------------------------------------------------------------------------
# CA-RequireMfaAllUsers (Tier 1)
# ---------------------------------------------------------------------------

function Get-CARequireMfaAllUsersSpec {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        Id       = 'CA-RequireMfaAllUsers'
        DisplayName = "$($script:CAPolicyNamePrefix)Require MFA for all users"
        Tier     = 1
        GrantControls = @{ Operator = 'OR'; BuiltInControls = @('mfa') }
        BuildConditions = {
            param($EmergencyGroupId)
            @{
                users = @{ includeUsers = @('All'); excludeGroups = @($EmergencyGroupId) }
                applications = @{ includeApplications = @('All') }
            }
        }
        OverlapPredicate = {
            param($Policy)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (@($Policy.Conditions.Users.IncludeUsers) -contains 'All')
        }
        ComplianceCheck = {
            param($Policy)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (@($Policy.Conditions.Users.IncludeUsers) -contains 'All') -and
                (@($Policy.Conditions.Applications.IncludeApplications) -contains 'All')
        }
    }
}

function Get-CA-RequireMfaAllUsersState {
    <#
    .SYNOPSIS
        Reads whether the "Require MFA for all users" report-only CA policy is compliant.
    .EXAMPLE
        Get-CA-RequireMfaAllUsersState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Get-BaselineCAControlState -Spec (Get-CARequireMfaAllUsersSpec)
}

function Set-CA-RequireMfaAllUsersState {
    <#
    .SYNOPSIS
        Idempotently creates/updates the "Require MFA for all users" report-only CA policy.
    .PARAMETER DesiredValue
        Ignored (always $true - "this control should exist and match its baseline definition").
    .PARAMETER CurrentValue
        Optional pre-fetched compliance bool.
    .PARAMETER Tier
        Ignored - this control is always Tier 1; present for BaselineCore's generic Tier pass-through.
    .PARAMETER ForceCreateDespiteOverlap
        Skip overlap detection for this control's creation.
    .EXAMPLE
        Set-CA-RequireMfaAllUsersState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [int]$Tier,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )
    Set-BaselineCAControlState -Spec (Get-CARequireMfaAllUsersSpec) -CurrentValue $CurrentValue -ForceCreateDespiteOverlap $ForceCreateDespiteOverlap
}

# ---------------------------------------------------------------------------
# CA-RequireMfaAdminRoles (Tier 1)
# ---------------------------------------------------------------------------

function Get-CARequireMfaAdminRolesSpec {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        Id       = 'CA-RequireMfaAdminRoles'
        DisplayName = "$($script:CAPolicyNamePrefix)Require MFA for admin roles"
        Tier     = 1
        GrantControls = @{ Operator = 'OR'; BuiltInControls = @('mfa') }
        BuildConditions = {
            param($EmergencyGroupId)
            @{
                users = @{ includeRoles = @(Get-BaselineCAAdminRoleTemplateIds); excludeGroups = @($EmergencyGroupId) }
                applications = @{ includeApplications = @('All') }
            }
        }
        OverlapPredicate = {
            param($Policy)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (@($Policy.Conditions.Users.IncludeRoles).Count -gt 0)
        }
        ComplianceCheck = {
            param($Policy)
            $expectedRoles = @(Get-BaselineCAAdminRoleTemplateIds | Sort-Object)
            $actualRoles = @($Policy.Conditions.Users.IncludeRoles | ForEach-Object { [string]$_ } | Sort-Object)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (-not (Compare-Object -ReferenceObject $expectedRoles -DifferenceObject $actualRoles)) -and
                (@($Policy.Conditions.Applications.IncludeApplications) -contains 'All')
        }
    }
}

function Get-CA-RequireMfaAdminRolesState {
    <#
    .SYNOPSIS
        Reads whether the "Require MFA for admin roles" report-only CA policy is compliant.
    .EXAMPLE
        Get-CA-RequireMfaAdminRolesState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Get-BaselineCAControlState -Spec (Get-CARequireMfaAdminRolesSpec)
}

function Set-CA-RequireMfaAdminRolesState {
    <#
    .SYNOPSIS
        Idempotently creates/updates the "Require MFA for admin roles" report-only CA policy.
    .PARAMETER DesiredValue
        Ignored (always $true).
    .PARAMETER CurrentValue
        Optional pre-fetched compliance bool.
    .PARAMETER Tier
        Ignored - always Tier 1.
    .PARAMETER ForceCreateDespiteOverlap
        Skip overlap detection for this control's creation.
    .EXAMPLE
        Set-CA-RequireMfaAdminRolesState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [int]$Tier,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )
    Set-BaselineCAControlState -Spec (Get-CARequireMfaAdminRolesSpec) -CurrentValue $CurrentValue -ForceCreateDespiteOverlap $ForceCreateDespiteOverlap
}

# ---------------------------------------------------------------------------
# CA-BlockLegacyAuth (Tier 1)
# ---------------------------------------------------------------------------

function Get-CABlockLegacyAuthSpec {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        Id       = 'CA-BlockLegacyAuth'
        DisplayName = "$($script:CAPolicyNamePrefix)Block legacy authentication"
        Tier     = 1
        GrantControls = @{ Operator = 'OR'; BuiltInControls = @('block') }
        BuildConditions = {
            param($EmergencyGroupId)
            @{
                users = @{ includeUsers = @('All'); excludeGroups = @($EmergencyGroupId) }
                applications = @{ includeApplications = @('All') }
                clientAppTypes = @('exchangeActiveSync', 'other')
            }
        }
        OverlapPredicate = {
            param($Policy)
            $clientAppTypes = @($Policy.Conditions.ClientAppTypes | ForEach-Object { [string]$_ })
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('block')) -and
                (($clientAppTypes -contains 'exchangeActiveSync') -or ($clientAppTypes -contains 'other'))
        }
        ComplianceCheck = {
            param($Policy)
            $clientAppTypes = @($Policy.Conditions.ClientAppTypes | ForEach-Object { [string]$_ } | Sort-Object)
            $expected = @('exchangeActiveSync', 'other' | Sort-Object)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('block')) -and
                (-not (Compare-Object -ReferenceObject $expected -DifferenceObject $clientAppTypes)) -and
                (@($Policy.Conditions.Users.IncludeUsers) -contains 'All') -and
                (@($Policy.Conditions.Applications.IncludeApplications) -contains 'All')
        }
    }
}

function Get-CA-BlockLegacyAuthState {
    <#
    .SYNOPSIS
        Reads whether the "Block legacy authentication" report-only CA policy is compliant.
    .EXAMPLE
        Get-CA-BlockLegacyAuthState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Get-BaselineCAControlState -Spec (Get-CABlockLegacyAuthSpec)
}

function Set-CA-BlockLegacyAuthState {
    <#
    .SYNOPSIS
        Idempotently creates/updates the "Block legacy authentication" report-only CA policy.
    .PARAMETER DesiredValue
        Ignored (always $true).
    .PARAMETER CurrentValue
        Optional pre-fetched compliance bool.
    .PARAMETER Tier
        Ignored - always Tier 1.
    .PARAMETER ForceCreateDespiteOverlap
        Skip overlap detection for this control's creation.
    .EXAMPLE
        Set-CA-BlockLegacyAuthState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [int]$Tier,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )
    Set-BaselineCAControlState -Spec (Get-CABlockLegacyAuthSpec) -CurrentValue $CurrentValue -ForceCreateDespiteOverlap $ForceCreateDespiteOverlap
}

# ---------------------------------------------------------------------------
# CA-RequireMfaAzureManagement (Tier 1)
# ---------------------------------------------------------------------------

function Get-CARequireMfaAzureManagementSpec {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        Id       = 'CA-RequireMfaAzureManagement'
        DisplayName = "$($script:CAPolicyNamePrefix)Require MFA for Azure management"
        Tier     = 1
        GrantControls = @{ Operator = 'OR'; BuiltInControls = @('mfa') }
        BuildConditions = {
            param($EmergencyGroupId)
            if (-not (Test-BaselineCAAzureManagementAppResolvable)) {
                throw "CA-RequireMfaAzureManagement: the well-known 'Microsoft Azure Management' application (appId $($script:CAAzureManagementAppId)) did not resolve via Get-MgServicePrincipal in this tenant - refusing to build a policy that targets an app that isn't there."
            }
            @{
                users = @{ includeUsers = @('All'); excludeGroups = @($EmergencyGroupId) }
                applications = @{ includeApplications = @($script:CAAzureManagementAppId) }
            }
        }
        OverlapPredicate = {
            param($Policy)
            # Matches either an existing policy that specifically targets the Azure
            # Management app, or one scoped to 'All' apps - a live-tenant check found
            # an existing "all users, all apps, MFA" policy that structurally passed
            # right by a narrower AppId-only check (Applications.IncludeApplications
            # was confirmed as {All}, which trivially already covers Azure Management
            # too), causing this control to create a genuinely redundant duplicate.
            $includedApps = @($Policy.Conditions.Applications.IncludeApplications)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (($includedApps -contains $script:CAAzureManagementAppId) -or ($includedApps -contains 'All'))
        }
        ComplianceCheck = {
            param($Policy)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (@($Policy.Conditions.Applications.IncludeApplications) -contains $script:CAAzureManagementAppId)
        }
    }
}

function Get-CA-RequireMfaAzureManagementState {
    <#
    .SYNOPSIS
        Reads whether the "Require MFA for Azure management" report-only CA policy is compliant.
    .EXAMPLE
        Get-CA-RequireMfaAzureManagementState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Get-BaselineCAControlState -Spec (Get-CARequireMfaAzureManagementSpec)
}

function Set-CA-RequireMfaAzureManagementState {
    <#
    .SYNOPSIS
        Idempotently creates/updates the "Require MFA for Azure management" report-only CA policy.
    .PARAMETER DesiredValue
        Ignored (always $true).
    .PARAMETER CurrentValue
        Optional pre-fetched compliance bool.
    .PARAMETER Tier
        Ignored - always Tier 1.
    .PARAMETER ForceCreateDespiteOverlap
        Skip overlap detection for this control's creation.
    .EXAMPLE
        Set-CA-RequireMfaAzureManagementState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [int]$Tier,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )
    Set-BaselineCAControlState -Spec (Get-CARequireMfaAzureManagementSpec) -CurrentValue $CurrentValue -ForceCreateDespiteOverlap $ForceCreateDespiteOverlap
}

# ---------------------------------------------------------------------------
# CA-RequireMfaSecurityInfoRegistration (Tier 1)
# ---------------------------------------------------------------------------

function Get-CARequireMfaSecurityInfoRegistrationSpec {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        Id       = 'CA-RequireMfaSecurityInfoRegistration'
        DisplayName = "$($script:CAPolicyNamePrefix)Require MFA to register security info"
        Tier     = 1
        GrantControls = @{ Operator = 'OR'; BuiltInControls = @('mfa') }
        BuildConditions = {
            param($EmergencyGroupId)
            @{
                users = @{ includeUsers = @('All'); excludeGroups = @($EmergencyGroupId) }
                applications = @{ includeUserActions = @('urn:user:registersecurityinfo') }
            }
        }
        OverlapPredicate = {
            param($Policy)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (@($Policy.Conditions.Applications.IncludeUserActions) -contains 'urn:user:registersecurityinfo')
        }
        ComplianceCheck = {
            param($Policy)
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (@($Policy.Conditions.Applications.IncludeUserActions) -contains 'urn:user:registersecurityinfo')
        }
    }
}

function Get-CA-RequireMfaSecurityInfoRegistrationState {
    <#
    .SYNOPSIS
        Reads whether the "Require MFA to register security info" report-only CA policy is compliant.
    .EXAMPLE
        Get-CA-RequireMfaSecurityInfoRegistrationState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Get-BaselineCAControlState -Spec (Get-CARequireMfaSecurityInfoRegistrationSpec)
}

function Set-CA-RequireMfaSecurityInfoRegistrationState {
    <#
    .SYNOPSIS
        Idempotently creates/updates the "Require MFA to register security info" report-only CA policy.
    .PARAMETER DesiredValue
        Ignored (always $true).
    .PARAMETER CurrentValue
        Optional pre-fetched compliance bool.
    .PARAMETER Tier
        Ignored - always Tier 1.
    .PARAMETER ForceCreateDespiteOverlap
        Skip overlap detection for this control's creation.
    .EXAMPLE
        Set-CA-RequireMfaSecurityInfoRegistrationState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [int]$Tier,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )
    Set-BaselineCAControlState -Spec (Get-CARequireMfaSecurityInfoRegistrationSpec) -CurrentValue $CurrentValue -ForceCreateDespiteOverlap $ForceCreateDespiteOverlap
}

# ---------------------------------------------------------------------------
# CA-RequireMfaGuestAccess (Tier 1)
# ---------------------------------------------------------------------------

function Get-CARequireMfaGuestAccessSpec {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        Id       = 'CA-RequireMfaGuestAccess'
        DisplayName = "$($script:CAPolicyNamePrefix)Require MFA for guest and external users"
        Tier     = 1
        GrantControls = @{ Operator = 'OR'; BuiltInControls = @('mfa') }
        BuildConditions = {
            param($EmergencyGroupId)
            @{
                users = @{
                    includeGuestsOrExternalUsers = @{
                        guestOrExternalUserTypes = 'internalGuest,b2bCollaborationGuest,b2bCollaborationMember,b2bDirectConnectUser,otherExternalUser,serviceProvider'
                        externalTenants = @{ membershipKind = 'all' }
                    }
                    excludeGroups = @($EmergencyGroupId)
                }
                applications = @{ includeApplications = @('All') }
            }
        }
        OverlapPredicate = {
            param($Policy)
            # Matches either an existing policy specifically scoped to guests, or one
            # scoped to 'All' users - a live-tenant check found an existing "all users,
            # all apps, MFA" policy that structurally passed right by a
            # guests-only-condition check, since 'All' users trivially already includes
            # guests, causing this control to create a genuinely redundant duplicate.
            $guestTypes = [string]$Policy.Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes
            $legacyGuest = @($Policy.Conditions.Users.IncludeUsers) -contains 'GuestsOrExternalUsers'
            $allUsers = @($Policy.Conditions.Users.IncludeUsers) -contains 'All'
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                ((-not [string]::IsNullOrWhiteSpace($guestTypes)) -or $legacyGuest -or $allUsers)
        }
        ComplianceCheck = {
            param($Policy)
            $guestTypes = [string]$Policy.Conditions.Users.IncludeGuestsOrExternalUsers.GuestOrExternalUserTypes
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (-not [string]::IsNullOrWhiteSpace($guestTypes)) -and
                (@($Policy.Conditions.Applications.IncludeApplications) -contains 'All')
        }
    }
}

function Get-CA-RequireMfaGuestAccessState {
    <#
    .SYNOPSIS
        Reads whether the "Require MFA for guest and external users" report-only CA policy is compliant.
    .EXAMPLE
        Get-CA-RequireMfaGuestAccessState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Get-BaselineCAControlState -Spec (Get-CARequireMfaGuestAccessSpec)
}

function Set-CA-RequireMfaGuestAccessState {
    <#
    .SYNOPSIS
        Idempotently creates/updates the "Require MFA for guest and external users" report-only CA policy.
    .PARAMETER DesiredValue
        Ignored (always $true).
    .PARAMETER CurrentValue
        Optional pre-fetched compliance bool.
    .PARAMETER Tier
        Ignored - always Tier 1.
    .PARAMETER ForceCreateDespiteOverlap
        Skip overlap detection for this control's creation.
    .EXAMPLE
        Set-CA-RequireMfaGuestAccessState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [int]$Tier,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )
    Set-BaselineCAControlState -Spec (Get-CARequireMfaGuestAccessSpec) -CurrentValue $CurrentValue -ForceCreateDespiteOverlap $ForceCreateDespiteOverlap
}

# ---------------------------------------------------------------------------
# CA-RequireMfaSignInRisk (Tier 2)
# ---------------------------------------------------------------------------

function Get-CARequireMfaSignInRiskSpec {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        Id       = 'CA-RequireMfaSignInRisk'
        DisplayName = "$($script:CAPolicyNamePrefix)Require MFA for medium and high sign-in risk"
        Tier     = 2
        GrantControls = @{ Operator = 'OR'; BuiltInControls = @('mfa') }
        BuildConditions = {
            param($EmergencyGroupId)
            @{
                users = @{ includeUsers = @('All'); excludeGroups = @($EmergencyGroupId) }
                applications = @{ includeApplications = @('All') }
                signInRiskLevels = @('medium', 'high')
            }
        }
        OverlapPredicate = {
            param($Policy)
            $riskLevels = @($Policy.Conditions.SignInRiskLevels | ForEach-Object { [string]$_ })
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and ($riskLevels.Count -gt 0)
        }
        ComplianceCheck = {
            param($Policy)
            $riskLevels = @($Policy.Conditions.SignInRiskLevels | ForEach-Object { [string]$_ } | Sort-Object)
            $expected = @('high', 'medium')
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa')) -and
                (-not (Compare-Object -ReferenceObject $expected -DifferenceObject $riskLevels)) -and
                (@($Policy.Conditions.Users.IncludeUsers) -contains 'All') -and
                (@($Policy.Conditions.Applications.IncludeApplications) -contains 'All')
        }
    }
}

function Get-CA-RequireMfaSignInRiskState {
    <#
    .SYNOPSIS
        Reads whether the Tier 2 "Require MFA for medium and high sign-in risk" report-only CA policy is compliant.
    .EXAMPLE
        Get-CA-RequireMfaSignInRiskState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Get-BaselineCAControlState -Spec (Get-CARequireMfaSignInRiskSpec)
}

function Set-CA-RequireMfaSignInRiskState {
    <#
    .SYNOPSIS
        Idempotently creates/updates the Tier 2 "Require MFA for medium and high sign-in risk" report-only CA policy.
    .PARAMETER DesiredValue
        Ignored (always $true).
    .PARAMETER CurrentValue
        Optional pre-fetched compliance bool.
    .PARAMETER Tier
        Ignored - always Tier 2.
    .PARAMETER ForceCreateDespiteOverlap
        Skip overlap detection for this control's creation.
    .EXAMPLE
        Set-CA-RequireMfaSignInRiskState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [int]$Tier,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )
    Set-BaselineCAControlState -Spec (Get-CARequireMfaSignInRiskSpec) -CurrentValue $CurrentValue -ForceCreateDespiteOverlap $ForceCreateDespiteOverlap
}

# ---------------------------------------------------------------------------
# CA-RequirePasswordChangeUserRisk (Tier 2)
# ---------------------------------------------------------------------------

function Get-CARequirePasswordChangeUserRiskSpec {
    [CmdletBinding()]
    [OutputType([hashtable])]
    param()
    return @{
        Id       = 'CA-RequirePasswordChangeUserRisk'
        DisplayName = "$($script:CAPolicyNamePrefix)Require password change for high user risk"
        Tier     = 2
        # Report-only means neither control is ever actually enforced; MFA is
        # included alongside passwordChange because Microsoft's own documented
        # high-user-risk baseline policy requires both together (a forced
        # password change with no re-authentication requirement is meaningless).
        GrantControls = @{ Operator = 'AND'; BuiltInControls = @('mfa', 'passwordChange') }
        BuildConditions = {
            param($EmergencyGroupId)
            @{
                users = @{ includeUsers = @('All'); excludeGroups = @($EmergencyGroupId) }
                applications = @{ includeApplications = @('All') }
                userRiskLevels = @('high')
            }
        }
        OverlapPredicate = {
            param($Policy)
            $riskLevels = @($Policy.Conditions.UserRiskLevels | ForEach-Object { [string]$_ })
            $grantControls = @($Policy.GrantControls.BuiltInControls | ForEach-Object { [string]$_ })
            ($grantControls -contains 'passwordChange') -and ($riskLevels -contains 'high')
        }
        ComplianceCheck = {
            param($Policy)
            $riskLevels = @($Policy.Conditions.UserRiskLevels | ForEach-Object { [string]$_ })
            (Test-BaselineCAGrantControlsMatch -Policy $Policy -ExpectedControls @('mfa', 'passwordChange') -ExpectedOperator 'AND') -and
                ($riskLevels -contains 'high') -and
                (@($Policy.Conditions.Users.IncludeUsers) -contains 'All') -and
                (@($Policy.Conditions.Applications.IncludeApplications) -contains 'All')
        }
    }
}

function Get-CA-RequirePasswordChangeUserRiskState {
    <#
    .SYNOPSIS
        Reads whether the Tier 2 "Require password change for high user risk" report-only CA policy is compliant.
    .DESCRIPTION
        Report-only means this policy, even once created, never actually forces
        a password change - it only logs what would have happened.
    .EXAMPLE
        Get-CA-RequirePasswordChangeUserRiskState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Get-BaselineCAControlState -Spec (Get-CARequirePasswordChangeUserRiskSpec)
}

function Set-CA-RequirePasswordChangeUserRiskState {
    <#
    .SYNOPSIS
        Idempotently creates/updates the Tier 2 "Require password change for high user risk" report-only CA policy.
    .PARAMETER DesiredValue
        Ignored (always $true).
    .PARAMETER CurrentValue
        Optional pre-fetched compliance bool.
    .PARAMETER Tier
        Ignored - always Tier 2.
    .PARAMETER ForceCreateDespiteOverlap
        Skip overlap detection for this control's creation.
    .EXAMPLE
        Set-CA-RequirePasswordChangeUserRiskState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [int]$Tier,

        [Parameter()]
        [bool]$ForceCreateDespiteOverlap = $false
    )
    Set-BaselineCAControlState -Spec (Get-CARequirePasswordChangeUserRiskSpec) -CurrentValue $CurrentValue -ForceCreateDespiteOverlap $ForceCreateDespiteOverlap
}

Export-ModuleMember -Function @(
    # Get-BaselineCAControlState/Set-BaselineCAControlState (the shared engine
    # every control below delegates to) are deliberately NOT exported: their
    # names end in "State" and would otherwise be misidentified as orphaned
    # catalog controls by Get-BaselineControlCatalog's Get-/Set-*State
    # auto-discovery scan, the same class of bug documented on that scan's
    # $ourModuleNames filter in BaselineCore.psm1. Get-Command (which that scan
    # uses) only sees a module's exported functions, so simply not exporting
    # them keeps them usable internally without tripping the scan. Use
    # InModuleScope 'ConditionalAccessControls' in tests to reach them directly.
    'Test-BaselineCATierAvailable'
    'Test-BaselineCASecurityDefaultsEnabled'
    'Get-BaselineCAPolicyListCache'
    'Find-BaselineCAPolicyByName'
    'Find-BaselineCAOverlap'
    'Get-BaselineCAEmergencyAccessGroupId'
    'Get-BaselineCAAdminRoleTemplateIds'
    'Test-BaselineCAAzureManagementAppResolvable'
    'Test-BaselineCAGrantControlsMatch'
    'Test-BaselineCAEmergencyGroupExcluded'
    'Get-BaselineConditionalAccessSummary'
    'Get-CA-RequireMfaAllUsersState', 'Set-CA-RequireMfaAllUsersState'
    'Get-CA-RequireMfaAdminRolesState', 'Set-CA-RequireMfaAdminRolesState'
    'Get-CA-BlockLegacyAuthState', 'Set-CA-BlockLegacyAuthState'
    'Get-CA-RequireMfaAzureManagementState', 'Set-CA-RequireMfaAzureManagementState'
    'Get-CA-RequireMfaSecurityInfoRegistrationState', 'Set-CA-RequireMfaSecurityInfoRegistrationState'
    'Get-CA-RequireMfaGuestAccessState', 'Set-CA-RequireMfaGuestAccessState'
    'Get-CA-RequireMfaSignInRiskState', 'Set-CA-RequireMfaSignInRiskState'
    'Get-CA-RequirePasswordChangeUserRiskState', 'Set-CA-RequirePasswordChangeUserRiskState'
)
