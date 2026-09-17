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
    return [pscustomobject]@{ Id = 'EntraID-GlobalAdminCount'; Value = $members.Count; Detail = ($members.Id -join ', ') }
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
        Voice call authentication methods.
    .EXAMPLE
        Get-EntraID-AuthMethodsHardeningState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $authenticator = Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -AuthenticationMethodConfigurationId 'MicrosoftAuthenticator' -ErrorAction Stop
    $sms = Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -AuthenticationMethodConfigurationId 'Sms' -ErrorAction Stop
    $voice = Get-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration -AuthenticationMethodConfigurationId 'Voice' -ErrorAction Stop

    $value = [pscustomobject]@{
        authenticatorEnabled = ([string]$authenticator.State -eq 'enabled')
        smsEnabled           = ([string]$sms.State -eq 'enabled')
        voiceEnabled          = ([string]$voice.State -eq 'enabled')
    }
    return [pscustomobject]@{ Id = 'EntraID-AuthMethodsHardening'; Value = $value }
}

function Set-EntraID-AuthMethodsHardeningState {
    <#
    .SYNOPSIS
        Idempotently enables/disables the Microsoft Authenticator, SMS, and Voice
        call authentication methods.
    .PARAMETER DesiredValue
        Object: { authenticatorEnabled, smsEnabled, voiceEnabled } (all bool).
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-EntraID-AuthMethodsHardeningState -DesiredValue ([pscustomobject]@{authenticatorEnabled=$true;smsEnabled=$false;voiceEnabled=$false})
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
    return [pscustomobject]@{ Id = 'EntraID-AuthMethodsHardening'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated MicrosoftAuthenticator/Sms/Voice method states.' }
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
    'Get-EntraID-RestrictAdminPortalAccessState', 'Set-EntraID-RestrictAdminPortalAccessState'
    'Get-EntraID-AuthMethodsHardeningState', 'Set-EntraID-AuthMethodsHardeningState'
    'Get-EntraID-MfaRegistrationCampaignState', 'Set-EntraID-MfaRegistrationCampaignState'
    'Get-EntraID-AdminPasswordResetNotificationState', 'Set-EntraID-AdminPasswordResetNotificationState'
)
