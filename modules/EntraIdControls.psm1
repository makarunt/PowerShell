#Requires -Version 7.0
<#
    EntraIdControls.psm1

    Get-/Set- function pairs for every EntraID control in the baseline inventory.
    Requires Microsoft.Graph (v2+) to be connected before use (Connect-BaselineWorkload
    -Connection Graph), except EntraID-UnifiedAuditLog, which uses an Exchange Online
    cmdlet (see BaselineCore's ControlConnectionOverrides).

    NOTE ON SCHEMA STABILITY: the request bodies used by
    Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration and
    Update-MgPolicyAuthenticationMethodPolicy (used by EntraID-AuthMethodsHardening and
    EntraID-MfaRegistrationCampaign) have changed shape between Microsoft.Graph module
    versions in the past. Validate the parameter/body shape below against the
    Microsoft.Graph.Identity.SignIns module version you have installed before relying
    on these two controls in production.
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# EntraID-UnifiedAuditLog
# ---------------------------------------------------------------------------

function Get-EntraID-UnifiedAuditLogState {
    <#
    .SYNOPSIS
        Reads whether unified audit log ingestion is enabled.
    .DESCRIPTION
        Uses Get-AdminAuditLogConfig (Exchange Online PowerShell), which is the
        Microsoft-documented surface for this tenant-wide setting.
    .EXAMPLE
        Get-EntraID-UnifiedAuditLogState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $cfg = Get-AdminAuditLogConfig -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-UnifiedAuditLog'; Value = [bool]$cfg.UnifiedAuditLogIngestionEnabled }
}

function Set-EntraID-UnifiedAuditLogState {
    <#
    .SYNOPSIS
        Idempotently sets unified audit log ingestion to the desired value.
    .PARAMETER DesiredValue
        Boolean: true to enable ingestion.
    .PARAMETER CurrentValue
        Optional pre-fetched current value, to avoid a redundant read.
    .EXAMPLE
        Set-EntraID-UnifiedAuditLogState -DesiredValue $true
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [bool]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )

    $current = if ($null -ne $CurrentValue) { [bool]$CurrentValue } else { (Get-EntraID-UnifiedAuditLogState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-UnifiedAuditLog'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }

    Set-AdminAuditLogConfig -UnifiedAuditLogIngestionEnabled $DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-UnifiedAuditLog'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated UnifiedAuditLogIngestionEnabled.' }
}

# ---------------------------------------------------------------------------
# EntraID-GlobalAdminCount (audit-only; no safe automated remediation)
# ---------------------------------------------------------------------------

function Get-EntraID-GlobalAdminCountState {
    <#
    .SYNOPSIS
        Counts active Global Administrator role members.
    .EXAMPLE
        Get-EntraID-GlobalAdminCountState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()

    $role = Get-MgDirectoryRole -Filter "DisplayName eq 'Global Administrator'" -ErrorAction Stop
    if (-not $role) {
        throw "The Global Administrator directory role is not activated in this tenant (Get-MgDirectoryRole returned no match)."
    }
    $members = @(Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All -ErrorAction Stop)
    # Member-enumeration ($members.Id) throws under Set-StrictMode -Version
    # Latest when $members is a genuinely empty array (confirmed directly) -
    # a legitimate shape here (a role with zero current members). ForEach-Object
    # never touches .Id at all when there's nothing to iterate.
    $memberIds = @($members | ForEach-Object { [string]$_.Id })
    return [pscustomobject]@{ Id = 'EntraID-GlobalAdminCount'; Value = $members.Count; Detail = ($memberIds -join ', ') }
}

function Set-EntraID-GlobalAdminCountState {
    <#
    .SYNOPSIS
        Not automatable: Global Administrator membership changes require human
        judgment and are never applied automatically. Always returns Skipped-Manual.
    .PARAMETER DesiredValue
        Ignored - present only to match the common Set- signature.
    .PARAMETER CurrentValue
        The current member count, echoed back for the log/report.
    .EXAMPLE
        Set-EntraID-GlobalAdminCountState -DesiredValue @{min=2;max=4} -CurrentValue 6
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )

    return [pscustomobject]@{
        Id            = 'EntraID-GlobalAdminCount'
        Status        = 'Skipped-Manual'
        PreviousValue = $CurrentValue
        AppliedValue  = $null
        Message       = 'No safe automated remediation for Global Administrator role membership. Review and adjust manually: Entra admin center > Identity > Roles & administrators > Global Administrator.'
    }
}

# ---------------------------------------------------------------------------
# EntraID-GuestInviteRestriction
# ---------------------------------------------------------------------------

function Get-EntraID-GuestInviteRestrictionState {
    <#
    .SYNOPSIS
        Reads who is allowed to invite guest users.
    .EXAMPLE
        Get-EntraID-GuestInviteRestrictionState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-GuestInviteRestriction'; Value = [string]$policy.AllowInvitesFrom }
}

function Set-EntraID-GuestInviteRestrictionState {
    <#
    .SYNOPSIS
        Idempotently sets who is allowed to invite guest users.
    .PARAMETER DesiredValue
        String enum understood by Update-MgPolicyAuthorizationPolicy -AllowInvitesFrom.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-GuestInviteRestrictionState -DesiredValue 'adminsAndGuestInviters'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { [string]$CurrentValue } else { (Get-EntraID-GuestInviteRestrictionState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-GuestInviteRestriction'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    # authorizationPolicy is a singleton resource (fixed path, no id in the URL), so
    # Update-MgPolicyAuthorizationPolicy takes no Id parameter at all - only
    # -BodyParameter is reliable across SDK versions (some versions also expose
    # -AllowInvitesFrom directly, but it isn't present in every installed version).
    Update-MgPolicyAuthorizationPolicy -BodyParameter @{ allowInvitesFrom = $DesiredValue } -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-GuestInviteRestriction'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated AllowInvitesFrom.' }
}

# ---------------------------------------------------------------------------
# EntraID-GuestUserRoleRestriction
# ---------------------------------------------------------------------------

function Get-EntraID-GuestUserRoleRestrictionState {
    <#
    .SYNOPSIS
        Reads the directory role id applied to guest users by default.
    .EXAMPLE
        Get-EntraID-GuestUserRoleRestrictionState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-GuestUserRoleRestriction'; Value = [string]$policy.GuestUserRoleId }
}

function Set-EntraID-GuestUserRoleRestrictionState {
    <#
    .SYNOPSIS
        Idempotently sets the default guest user role.
    .PARAMETER DesiredValue
        Directory role template id string, e.g. the Restricted Guest User role id.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-GuestUserRoleRestrictionState -DesiredValue '2af84b1e-32c8-42b7-82bc-daa82404023b'
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [string]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { [string]$CurrentValue } else { (Get-EntraID-GuestUserRoleRestrictionState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-GuestUserRoleRestriction'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    # See EntraID-GuestInviteRestriction's Set- function: authorizationPolicy is a
    # singleton, so no Id parameter exists on Update-MgPolicyAuthorizationPolicy.
    Update-MgPolicyAuthorizationPolicy -BodyParameter @{ guestUserRoleId = $DesiredValue } -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-GuestUserRoleRestriction'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated GuestUserRoleId.' }
}

# ---------------------------------------------------------------------------
# EntraID-BlockUserConsentToApps
# ---------------------------------------------------------------------------

function Get-EntraID-BlockUserConsentToAppsState {
    <#
    .SYNOPSIS
        Reads the permission grant policies assigned to non-admin users for app consent.
    .EXAMPLE
        Get-EntraID-BlockUserConsentToAppsState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop
    $assigned = @($policy.DefaultUserRolePermissions.PermissionGrantPoliciesAssigned)
    return [pscustomobject]@{ Id = 'EntraID-BlockUserConsentToApps'; Value = [pscustomobject]@{ permissionGrantPoliciesAssigned = $assigned } }
}

function Set-EntraID-BlockUserConsentToAppsState {
    <#
    .SYNOPSIS
        Idempotently sets the permission grant policies assigned for user app consent.
    .PARAMETER DesiredValue
        Object: { permissionGrantPoliciesAssigned: [] }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-BlockUserConsentToAppsState -DesiredValue ([pscustomobject]@{permissionGrantPoliciesAssigned=@()})
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-EntraID-BlockUserConsentToAppsState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-BlockUserConsentToApps'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    $body = @{ defaultUserRolePermissions = @{ permissionGrantPoliciesAssigned = @($DesiredValue.permissionGrantPoliciesAssigned) } }
    Update-MgPolicyAuthorizationPolicy -BodyParameter $body -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-BlockUserConsentToApps'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated DefaultUserRolePermissions.PermissionGrantPoliciesAssigned.' }
}

# ---------------------------------------------------------------------------
# EntraID-BlockSelfServiceAppCreation
# ---------------------------------------------------------------------------

function Get-EntraID-BlockSelfServiceAppCreationState {
    <#
    .SYNOPSIS
        Reads whether non-admin users can register app registrations or create tenants.
    .DESCRIPTION
        DEVIATION FROM THE V2 GAP-ANALYSIS SPEC, FLAGGED DELIBERATELY: the spec
        that produced this v2 asked to add allowedToCreateSecurityGroups here,
        alongside allowedToCreateApps/allowedToCreateTenants, in the same
        Update-MgPolicyAuthorizationPolicy call. This toolkit already has a
        separate, working control for that exact field -
        EntraID-BlockSelfServiceSecurityGroupCreation, below - built in an
        earlier iteration of this toolkit and carried into v2 unchanged.
        Literally merging the field in here as asked would leave TWO controls
        independently PATCHing sibling properties of the same
        defaultUserRolePermissions sub-object within one Apply run. That is
        exactly the failure class root-caused (though never fully resolved)
        during this project's app-only-authentication work: of several
        sequential writes to different sub-fields of the authorizationPolicy
        singleton in one run, only the first and last were observed to persist
        on a real tenant - the middle ones silently reverted despite each
        individually reporting success. Interactive auth has not reproduced
        that specific bug in this project, but there is no reason to
        reintroduce the exact write pattern that caused it when a working
        single-owner control for this field already exists. Resolution: keep
        EntraID-BlockSelfServiceSecurityGroupCreation as the sole owner of
        allowedToCreateSecurityGroups; this control keeps owning only
        allowedToCreateApps/allowedToCreateTenants, as in v1.
    .EXAMPLE
        Get-EntraID-BlockSelfServiceAppCreationState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop
    $value = [pscustomobject]@{
        allowedToCreateApps    = [bool]$policy.DefaultUserRolePermissions.AllowedToCreateApps
        allowedToCreateTenants = [bool]$policy.DefaultUserRolePermissions.AllowedToCreateTenants
    }
    return [pscustomobject]@{ Id = 'EntraID-BlockSelfServiceAppCreation'; Value = $value }
}

function Set-EntraID-BlockSelfServiceAppCreationState {
    <#
    .SYNOPSIS
        Idempotently blocks/unblocks self-service app registration and tenant creation.
    .PARAMETER DesiredValue
        Object: { allowedToCreateApps: bool, allowedToCreateTenants: bool }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-BlockSelfServiceAppCreationState -DesiredValue ([pscustomobject]@{allowedToCreateApps=$false;allowedToCreateTenants=$false})
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-EntraID-BlockSelfServiceAppCreationState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-BlockSelfServiceAppCreation'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    $body = @{
        defaultUserRolePermissions = @{
            allowedToCreateApps    = [bool]$DesiredValue.allowedToCreateApps
            allowedToCreateTenants = [bool]$DesiredValue.allowedToCreateTenants
        }
    }
    Update-MgPolicyAuthorizationPolicy -BodyParameter $body -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-BlockSelfServiceAppCreation'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated AllowedToCreateApps/AllowedToCreateTenants.' }
}

# ---------------------------------------------------------------------------
# EntraID-BlockSelfServiceSecurityGroupCreation
# ---------------------------------------------------------------------------

function Get-EntraID-BlockSelfServiceSecurityGroupCreationState {
    <#
    .SYNOPSIS
        Reads whether non-admin users can create security groups.
    .EXAMPLE
        Get-EntraID-BlockSelfServiceSecurityGroupCreationState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-MgPolicyAuthorizationPolicy -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-BlockSelfServiceSecurityGroupCreation'; Value = [bool]$policy.DefaultUserRolePermissions.AllowedToCreateSecurityGroups }
}

function Set-EntraID-BlockSelfServiceSecurityGroupCreationState {
    <#
    .SYNOPSIS
        Idempotently blocks/unblocks self-service security group creation.
    .PARAMETER DesiredValue
        Boolean - $false to restrict security group creation to admins.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-BlockSelfServiceSecurityGroupCreationState -DesiredValue $false
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [bool]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { [bool]$CurrentValue } else { (Get-EntraID-BlockSelfServiceSecurityGroupCreationState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-BlockSelfServiceSecurityGroupCreation'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    # Same pattern as EntraID-BlockSelfServiceAppCreation: send only the one
    # defaultUserRolePermissions sub-property being changed. Confirmed on a
    # real tenant that this does NOT reset allowedToCreateApps/
    # allowedToCreateTenants when only they were previously PATCHed, so the
    # reverse (patching this field alone without touching those two) is safe
    # the same way - Graph merges within defaultUserRolePermissions rather
    # than replacing the whole nested object.
    $body = @{
        defaultUserRolePermissions = @{
            allowedToCreateSecurityGroups = $DesiredValue
        }
    }
    Update-MgPolicyAuthorizationPolicy -BodyParameter $body -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-BlockSelfServiceSecurityGroupCreation'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated AllowedToCreateSecurityGroups.' }
}

# ---------------------------------------------------------------------------
# EntraID-AdminConsentWorkflow
# ---------------------------------------------------------------------------

function Get-EntraID-AdminConsentWorkflowState {
    <#
    .SYNOPSIS
        Reads the admin consent request workflow policy.
    .EXAMPLE
        Get-EntraID-AdminConsentWorkflowState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-MgPolicyAdminConsentRequestPolicy -ErrorAction Stop
    $reviewers = @($policy.Reviewers | ForEach-Object {
        [pscustomobject]@{
            query     = [string]$_.Query
            queryType = [string]$_.QueryType
            queryRoot = if ($_.QueryRoot) { [string]$_.QueryRoot } else { $null }
        }
    })
    $value = [pscustomobject]@{
        isEnabled             = [bool]$policy.IsEnabled
        notifyReviewers       = [bool]$policy.NotifyReviewers
        remindersEnabled      = [bool]$policy.RemindersEnabled
        requestDurationInDays = [int]$policy.RequestDurationInDays
        reviewers             = $reviewers
    }
    return [pscustomobject]@{ Id = 'EntraID-AdminConsentWorkflow'; Value = $value }
}

function Set-EntraID-AdminConsentWorkflowState {
    <#
    .SYNOPSIS
        Idempotently configures the admin consent request workflow.
    .DESCRIPTION
        reviewers has no safe universal default - it must reference real user,
        group, or role IDs in this tenant. Config validation
        (Test-BaselineApplyReadiness's requiresPopulatedFields check in
        BaselineCore.psm1, wired for this control in
        config/baseline.config.json) refuses to run Apply at all while
        desiredValue.reviewers is empty - the same pattern already used for
        ExchangeOnline-DkimSigning's domain list and
        Teams-RestrictFederation's allowed-domains list. This function has its
        own defense-in-depth check for the same condition in case it's ever
        invoked directly (e.g. from Restore, which does not go through
        Test-BaselineApplyReadiness).
    .PARAMETER DesiredValue
        Object: { isEnabled, notifyReviewers, remindersEnabled: bool,
        requestDurationInDays: int, reviewers: [ { query, queryType, queryRoot? } ] }.
        Each reviewers entry is an accessReviewReviewerScope
        (query/queryType/optional queryRoot), e.g.
        { query = '/v1.0/users/<id>'; queryType = 'MicrosoftGraph' } for a
        specific user, or a group/role query for a broader scope.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-AdminConsentWorkflowState -DesiredValue ([pscustomobject]@{isEnabled=$true;notifyReviewers=$true;remindersEnabled=$true;requestDurationInDays=30;reviewers=@(@{query='/v1.0/users/00000000-0000-0000-0000-000000000000';queryType='MicrosoftGraph'})})
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-EntraID-AdminConsentWorkflowState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-AdminConsentWorkflow'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }

    $reviewers = @($DesiredValue.reviewers)
    if ($reviewers.Count -eq 0) {
        throw "EntraID-AdminConsentWorkflow requires at least one entry in desiredValue.reviewers (Microsoft Graph has no safe default reviewer) - update config/baseline.config.json with real user/group/role ids before running Apply."
    }
    $reviewerBodies = @($reviewers | ForEach-Object {
        $r = @{ query = [string]$_.query; queryType = [string]$_.queryType }
        if ($_.PSObject.Properties['queryRoot'] -and $_.queryRoot) { $r['queryRoot'] = [string]$_.queryRoot }
        $r
    })

    Update-MgPolicyAdminConsentRequestPolicy -IsEnabled:([bool]$DesiredValue.isEnabled) -NotifyReviewers:([bool]$DesiredValue.notifyReviewers) `
        -RemindersEnabled:([bool]$DesiredValue.remindersEnabled) -RequestDurationInDays ([int]$DesiredValue.requestDurationInDays) `
        -Reviewers $reviewerBodies -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-AdminConsentWorkflow'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated admin consent request policy.' }
}

# ---------------------------------------------------------------------------
# EntraID-GaNotLocalAdminOnJoin (audit-only; Preview feature, no stable v1.0 API)
# ---------------------------------------------------------------------------

function Get-EntraID-GaNotLocalAdminOnJoinState {
    <#
    .SYNOPSIS
        Audit-only placeholder: as of this writing, "Global Administrator role
        does not become a local administrator on newly Entra-joined devices" is
        a Preview-labeled device setting with no stable (non-beta) Microsoft
        Graph or PowerShell surface. Always reports Unknown.
    .DESCRIPTION
        Investigated against current Microsoft documentation before
        implementing: the setting lives on Entra ID's device registration
        policy, exposed today only through the Microsoft Graph beta endpoint
        (policies/deviceRegistrationPolicy, still Preview) rather than v1.0.
        This toolkit does not call beta Graph endpoints for automated
        remediation (no stability/support guarantee for a Preview feature), so
        this control follows the same audit-only pattern as
        EntraID-RestrictAdminPortalAccess: no Set- automation, Skipped-Manual
        with an exact GUI path. Revisit once Microsoft ships a v1.0 (non-beta)
        API for this setting.
    .EXAMPLE
        Get-EntraID-GaNotLocalAdminOnJoinState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    return [pscustomobject]@{
        Id     = 'EntraID-GaNotLocalAdminOnJoin'
        Value  = $null
        Detail = 'No stable (v1.0) Microsoft Graph or PowerShell API is documented for this Preview-labeled device setting as of this writing (only a beta Graph endpoint exists); verify manually in the Entra admin center.'
    }
}

function Set-EntraID-GaNotLocalAdminOnJoinState {
    <#
    .SYNOPSIS
        Not automatable: no stable (non-beta) API exists for this Preview
        feature. Always returns Skipped-Manual.
    .PARAMETER DesiredValue
        Ignored.
    .PARAMETER CurrentValue
        Echoed back for the log/report.
    .EXAMPLE
        Set-EntraID-GaNotLocalAdminOnJoinState -DesiredValue $true -CurrentValue $null
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    return [pscustomobject]@{
        Id            = 'EntraID-GaNotLocalAdminOnJoin'
        Status        = 'Skipped-Manual'
        PreviousValue = $CurrentValue
        AppliedValue  = $null
        Message       = 'No stable API exists for this Preview feature. Change manually: Entra admin center > Identity > Devices > Device settings > "Additional local administrators on Microsoft Entra joined devices" (ensure Global Administrator is not implicitly granted local admin).'
    }
}

# ---------------------------------------------------------------------------
# EntraID-RestrictAdminPortalAccess (audit-only; no confirmed cmdlet)
# ---------------------------------------------------------------------------

function Get-EntraID-RestrictAdminPortalAccessState {
    <#
    .SYNOPSIS
        Audit-only placeholder: there is no publicly documented, stable Graph
        property for this tenant setting as of this writing. Always reports Unknown.
    .EXAMPLE
        Get-EntraID-RestrictAdminPortalAccessState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    return [pscustomobject]@{
        Id     = 'EntraID-RestrictAdminPortalAccess'
        Value  = $null
        Detail = 'No stable Microsoft Graph property is documented for this setting as of this writing; verify manually in the Entra admin center.'
    }
}

function Set-EntraID-RestrictAdminPortalAccessState {
    <#
    .SYNOPSIS
        Not automatable: no confirmed cmdlet exists for this setting. Always
        returns Skipped-Manual.
    .PARAMETER DesiredValue
        Ignored.
    .PARAMETER CurrentValue
        Echoed back for the log/report.
    .EXAMPLE
        Set-EntraID-RestrictAdminPortalAccessState -DesiredValue $true -CurrentValue $null
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    return [pscustomobject]@{
        Id            = 'EntraID-RestrictAdminPortalAccess'
        Status        = 'Skipped-Manual'
        PreviousValue = $CurrentValue
        AppliedValue  = $null
        Message       = 'No confirmed cmdlet exists for this setting. Change manually: Entra admin center > Identity > Users > User settings > "Restrict access to Microsoft Entra admin center" = Yes.'
    }
}

# ---------------------------------------------------------------------------
# EntraID-AuthMethodsHardening
# ---------------------------------------------------------------------------

function Get-EntraID-AuthMethodsHardeningState {
    <#
    .SYNOPSIS
        Reads the enabled/disabled state of the Microsoft Authenticator, SMS, and
        Voice call authentication methods, plus the tenant-level
        systemCredentialPreferences ("system-preferred multifactor
        authentication") setting.
    .DESCRIPTION
        systemCredentialPreferences.state is read from the top-level
        authenticationMethodsPolicy resource (Get-MgPolicyAuthenticationMethodPolicy),
        not from the per-method AuthenticationMethodConfiguration cmdlet used for
        the other three methods - it isn't a method configuration, it's a
        separate tenant-wide preference on the policy itself. Microsoft has been
        gradually rolling out the sign-in-time *effect* of this setting
        tenant-by-tenant through roughly September 2026: a tenant can show this
        correctly configured (state = enabled) without yet visibly changing
        sign-in prompts. That is not a configuration error and this toolkit does
        not attempt to detect or "fix" it - see the README's "systemCredentialPreferences
        rollout" note. The Detail field below flags it so a report reader doesn't
        mistake it for a toolkit bug.
    .EXAMPLE
        Get-EntraID-AuthMethodsHardeningState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $authenticator = Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -AuthenticationMethodConfigurationId 'MicrosoftAuthenticator' -ErrorAction Stop
    $sms = Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -AuthenticationMethodConfigurationId 'Sms' -ErrorAction Stop
    $voice = Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -AuthenticationMethodConfigurationId 'Voice' -ErrorAction Stop
    $policy = Get-MgPolicyAuthenticationMethodPolicy -ErrorAction Stop

    # Confirmed against a real tenant: systemCredentialPreferences isn't
    # modeled as a direct typed property on every installed Microsoft.Graph
    # SDK version (it's a newer, still-rolling-out field - see this
    # function's own .DESCRIPTION) - $policy.SystemCredentialPreferences
    # throws "The property 'SystemCredentialPreferences' cannot be found on
    # this object" under Set-StrictMode -Version Latest when that's the case.
    # Get-BaselineGraphPropertyValue falls back to $policy.AdditionalProperties
    # instead of dotting in directly.
    $systemCredentialPreferences = Get-BaselineGraphPropertyValue -InputObject $policy -Name 'SystemCredentialPreferences'
    $systemCredentialPreferencesState = [string](Get-BaselineGraphPropertyValue -InputObject $systemCredentialPreferences -Name 'State')

    $value = [pscustomobject]@{
        authenticatorEnabled = ([string]$authenticator.State -eq 'enabled')
        smsEnabled           = ([string]$sms.State -eq 'enabled')
        voiceEnabled          = ([string]$voice.State -eq 'enabled')
        systemCredentialPreferences = [pscustomobject]@{ state = $systemCredentialPreferencesState }
    }
    return [pscustomobject]@{
        Id     = 'EntraID-AuthMethodsHardening'
        Value  = $value
        Detail = "Microsoft has been gradually rolling out systemCredentialPreferences's sign-in-time effect tenant-by-tenant through roughly September 2026 - a compliant reading here does not guarantee sign-in behavior has visibly changed yet."
    }
}

function Set-EntraID-AuthMethodsHardeningState {
    <#
    .SYNOPSIS
        Idempotently enables/disables the Microsoft Authenticator, SMS, and Voice
        call authentication methods, and sets the systemCredentialPreferences
        state.
    .PARAMETER DesiredValue
        Object: { authenticatorEnabled, smsEnabled, voiceEnabled: bool,
        systemCredentialPreferences: { state: 'enabled'|'disabled' } }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-AuthMethodsHardeningState -DesiredValue ([pscustomobject]@{authenticatorEnabled=$true;smsEnabled=$false;voiceEnabled=$false;systemCredentialPreferences=@{state='enabled'}})
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-EntraID-AuthMethodsHardeningState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-AuthMethodsHardening'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }

    $methodStates = @{
        'MicrosoftAuthenticator' = [bool]$DesiredValue.authenticatorEnabled
        'Sms'                    = [bool]$DesiredValue.smsEnabled
        'Voice'                  = [bool]$DesiredValue.voiceEnabled
    }
    foreach ($methodId in $methodStates.Keys) {
        $state = if ($methodStates[$methodId]) { 'enabled' } else { 'disabled' }
        Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -AuthenticationMethodConfigurationId $methodId -BodyParameter @{ '@odata.type' = "#microsoft.graph.$($methodId)AuthenticationMethodConfiguration"; state = $state } -ErrorAction Stop
    }
    Update-MgPolicyAuthenticationMethodPolicy -BodyParameter @{ systemCredentialPreferences = @{ state = [string]$DesiredValue.systemCredentialPreferences.state } } -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-AuthMethodsHardening'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated MicrosoftAuthenticator/Sms/Voice method states and systemCredentialPreferences.' }
}

# ---------------------------------------------------------------------------
# EntraID-MfaRegistrationCampaign
# ---------------------------------------------------------------------------

function Get-EntraID-MfaRegistrationCampaignState {
    <#
    .SYNOPSIS
        Reads the MFA registration campaign (nudge) state.
    .EXAMPLE
        Get-EntraID-MfaRegistrationCampaignState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-MgPolicyAuthenticationMethodPolicy -ErrorAction Stop
    $campaign = $policy.RegistrationEnforcement.AuthenticationMethodsRegistrationCampaign
    $includeTargets = @($campaign.IncludeTargets | ForEach-Object {
        [pscustomobject]@{
            targetType                  = [string]$_.TargetType
            id                          = [string]$_.Id
            targetedAuthenticationMethod = [string]$_.TargetedAuthenticationMethod
        }
    })
    $value = [pscustomobject]@{
        state                 = [string]$campaign.State
        snoozeDurationInDays  = [int]$campaign.SnoozeDurationInDays
        includeTargets        = $includeTargets
    }
    return [pscustomobject]@{ Id = 'EntraID-MfaRegistrationCampaign'; Value = $value }
}

function Set-EntraID-MfaRegistrationCampaignState {
    <#
    .SYNOPSIS
        Idempotently sets the MFA registration campaign state and snooze duration.
    .PARAMETER DesiredValue
        Object: { state: 'enabled'|'disabled', snoozeDurationInDays: int,
        includeTargets: [ { targetType, id, targetedAuthenticationMethod } ] }.
        includeTargets is required by the Graph API - the campaign has no effect
        without at least one target; the seed config uses the documented
        "all_users" special group id to target the whole tenant.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-MfaRegistrationCampaignState -DesiredValue ([pscustomobject]@{state='enabled';snoozeDurationInDays=1;includeTargets=@(@{targetType='group';id='all_users';targetedAuthenticationMethod='microsoftAuthenticator'})})
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-EntraID-MfaRegistrationCampaignState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'EntraID-MfaRegistrationCampaign'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    $includeTargets = @($DesiredValue.includeTargets | ForEach-Object {
        @{
            targetType                  = [string]$_.targetType
            id                          = [string]$_.id
            targetedAuthenticationMethod = [string]$_.targetedAuthenticationMethod
        }
    })
    if ($includeTargets.Count -eq 0) {
        throw "EntraID-MfaRegistrationCampaign requires at least one entry in desiredValue.includeTargets (the Graph API rejects an empty target list); update config/baseline.config.json before running Apply."
    }
    $body = @{
        registrationEnforcement = @{
            authenticationMethodsRegistrationCampaign = @{
                state                = [string]$DesiredValue.state
                snoozeDurationInDays = [int]$DesiredValue.snoozeDurationInDays
                includeTargets       = $includeTargets
            }
        }
    }
    Update-MgPolicyAuthenticationMethodPolicy -BodyParameter $body -ErrorAction Stop
    return [pscustomobject]@{ Id = 'EntraID-MfaRegistrationCampaign'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated RegistrationEnforcement.AuthenticationMethodsRegistrationCampaign.' }
}

# ---------------------------------------------------------------------------
# EntraID-AdminPasswordResetNotification (audit-only; no confirmed cmdlet)
# ---------------------------------------------------------------------------

function Get-EntraID-AdminPasswordResetNotificationState {
    <#
    .SYNOPSIS
        Audit-only placeholder: no publicly documented, stable Graph property
        exists for this setting as of this writing. Always reports Unknown.
    .EXAMPLE
        Get-EntraID-AdminPasswordResetNotificationState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    return [pscustomobject]@{
        Id     = 'EntraID-AdminPasswordResetNotification'
        Value  = $null
        Detail = 'No stable Microsoft Graph property is documented for this setting as of this writing; verify manually in the Entra admin center.'
    }
}

function Set-EntraID-AdminPasswordResetNotificationState {
    <#
    .SYNOPSIS
        Not automatable: no confirmed cmdlet exists for this setting. Always
        returns Skipped-Manual.
    .PARAMETER DesiredValue
        Ignored.
    .PARAMETER CurrentValue
        Echoed back for the log/report.
    .EXAMPLE
        Set-EntraID-AdminPasswordResetNotificationState -DesiredValue $true -CurrentValue $null
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [AllowNull()]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    return [pscustomobject]@{
        Id            = 'EntraID-AdminPasswordResetNotification'
        Status        = 'Skipped-Manual'
        PreviousValue = $CurrentValue
        AppliedValue  = $null
        Message       = 'No confirmed cmdlet exists for this setting. Change manually: Entra admin center > Identity > Users > User settings.'
    }
}

Export-ModuleMember -Function @(
    'Get-EntraID-UnifiedAuditLogState', 'Set-EntraID-UnifiedAuditLogState'
    'Get-EntraID-GlobalAdminCountState', 'Set-EntraID-GlobalAdminCountState'
    'Get-EntraID-GuestInviteRestrictionState', 'Set-EntraID-GuestInviteRestrictionState'
    'Get-EntraID-GuestUserRoleRestrictionState', 'Set-EntraID-GuestUserRoleRestrictionState'
    'Get-EntraID-BlockUserConsentToAppsState', 'Set-EntraID-BlockUserConsentToAppsState'
    'Get-EntraID-BlockSelfServiceAppCreationState', 'Set-EntraID-BlockSelfServiceAppCreationState'
    'Get-EntraID-BlockSelfServiceSecurityGroupCreationState', 'Set-EntraID-BlockSelfServiceSecurityGroupCreationState'
    'Get-EntraID-AdminConsentWorkflowState', 'Set-EntraID-AdminConsentWorkflowState'
    'Get-EntraID-GaNotLocalAdminOnJoinState', 'Set-EntraID-GaNotLocalAdminOnJoinState'
    'Get-EntraID-RestrictAdminPortalAccessState', 'Set-EntraID-RestrictAdminPortalAccessState'
    'Get-EntraID-AuthMethodsHardeningState', 'Set-EntraID-AuthMethodsHardeningState'
    'Get-EntraID-MfaRegistrationCampaignState', 'Set-EntraID-MfaRegistrationCampaignState'
    'Get-EntraID-AdminPasswordResetNotificationState', 'Set-EntraID-AdminPasswordResetNotificationState'
)
