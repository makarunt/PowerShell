#Requires -Version 7.0
<#
    TeamsControls.psm1

    Get-/Set- function pairs for every Teams control in the baseline inventory.
    Requires MicrosoftTeams (v6+) to be connected before use (Connect-BaselineWorkload
    -Connection Teams).
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Teams-BlockConsumerContact
# ---------------------------------------------------------------------------

function Get-Teams-BlockConsumerContactState {
    <#
    .SYNOPSIS
        Reads whether contact with unmanaged consumer Teams/Skype accounts is
        allowed, and whether federation with trial-only tenants is allowed.
    .EXAMPLE
        Get-Teams-BlockConsumerContactState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $cfg = Get-CsTenantFederationConfiguration -ErrorAction Stop
    $value = [pscustomobject]@{
        allowTeamsConsumer         = [bool]$cfg.AllowTeamsConsumer
        allowTeamsConsumerInbound  = [bool]$cfg.AllowTeamsConsumerInbound
        externalAccessWithTrialTenants = [string]$cfg.ExternalAccessWithTrialTenants
    }
    return [pscustomobject]@{ Id = 'Teams-BlockConsumerContact'; Value = $value }
}

function Set-Teams-BlockConsumerContactState {
    <#
    .SYNOPSIS
        Idempotently blocks/allows contact with unmanaged consumer Teams/Skype
        accounts and federation with trial-only tenants.
    .DESCRIPTION
        Microsoft made "Blocked" the tenant-wide default for
        ExternalAccessWithTrialTenants starting July 29, 2024, so on many
        tenants this may already read compliant before this toolkit ever runs;
        it's still asserted explicitly here rather than relying on the
        inherited default, since a tenant provisioned before that date - or
        one where an admin changed it - may not have it.
    .PARAMETER DesiredValue
        Object: { allowTeamsConsumer, allowTeamsConsumerInbound: bool,
        externalAccessWithTrialTenants: 'Allowed'|'Blocked' }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-Teams-BlockConsumerContactState -DesiredValue ([pscustomobject]@{allowTeamsConsumer=$false;allowTeamsConsumerInbound=$false;externalAccessWithTrialTenants='Blocked'})
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
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-Teams-BlockConsumerContactState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'Teams-BlockConsumerContact'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-CsTenantFederationConfiguration -AllowTeamsConsumer:([bool]$DesiredValue.allowTeamsConsumer) -AllowTeamsConsumerInbound:([bool]$DesiredValue.allowTeamsConsumerInbound) `
        -ExternalAccessWithTrialTenants ([string]$DesiredValue.externalAccessWithTrialTenants) -ErrorAction Stop
    return [pscustomobject]@{ Id = 'Teams-BlockConsumerContact'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated CsTenantFederationConfiguration consumer contact and trial-tenant settings.' }
}

# ---------------------------------------------------------------------------
# Teams-RestrictFederation
# ---------------------------------------------------------------------------

function Get-Teams-RestrictFederationState {
    <#
    .SYNOPSIS
        Reads the external Teams federation mode and allowed domain list.
    .DESCRIPTION
        Models the tenant federation configuration as { mode, allowedDomains }:
        mode is 'AllowAll' when AllowFederatedUsers is true with no domain
        restriction, 'AllowSpecific' when an allow-list of domains is configured,
        or 'BlockAll' when AllowFederatedUsers is false.
    .EXAMPLE
        Get-Teams-RestrictFederationState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $cfg = Get-CsTenantFederationConfiguration -ErrorAction Stop
    $allowedDomains = @()
    if ($cfg.AllowedDomains -and $cfg.AllowedDomains.PSObject.Properties['AllowedDomain']) {
        $allowedDomains = @($cfg.AllowedDomains.AllowedDomain | ForEach-Object { [string]$_.Domain })
    }

    $mode = if (-not $cfg.AllowFederatedUsers) { 'BlockAll' }
            elseif ($allowedDomains.Count -gt 0) { 'AllowSpecific' }
            else { 'AllowAll' }

    $value = [pscustomobject]@{ mode = $mode; allowedDomains = $allowedDomains }
    return [pscustomobject]@{ Id = 'Teams-RestrictFederation'; Value = $value }
}

function Set-Teams-RestrictFederationState {
    <#
    .SYNOPSIS
        Idempotently sets external Teams federation mode and allowed domains.
    .DESCRIPTION
        An empty allowedDomains list combined with mode 'AllowSpecific' is a valid
        but severe configuration meaning "block federation with every external
        domain." To avoid applying that silently, this function requires either a
        non-empty allowedDomains list or an explicit -AcknowledgeRisk switch when
        mode is 'AllowSpecific' and the list is empty.
    .PARAMETER DesiredValue
        Object: { mode: 'AllowAll'|'AllowSpecific'|'BlockAll', allowedDomains: [...] }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .PARAMETER AcknowledgeRisk
        Must be set to proceed when mode is 'AllowSpecific' and allowedDomains is empty.
    .EXAMPLE
        Set-Teams-RestrictFederationState -DesiredValue ([pscustomobject]@{mode='AllowSpecific';allowedDomains=@('partner.com')})
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [object]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue,

        [Parameter()]
        [switch]$AcknowledgeRisk
    )

    $mode = [string]$DesiredValue.mode
    $allowedDomains = @($DesiredValue.allowedDomains)

    if ($mode -eq 'AllowSpecific' -and $allowedDomains.Count -eq 0 -and -not $AcknowledgeRisk) {
        throw "Teams-RestrictFederation: desiredValue.allowedDomains is empty while mode is 'AllowSpecific'. This blocks federation with EVERY external domain. This is a valid but severe configuration - re-run Apply with -AcknowledgeFederationBlockAll to confirm this is intentional, or populate config/baseline.config.json controls[].desiredValue.allowedDomains with the partner domains that should remain allowed."
    }

    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-Teams-RestrictFederationState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'Teams-RestrictFederation'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }

    switch ($mode) {
        'BlockAll' {
            Set-CsTenantFederationConfiguration -AllowFederatedUsers:$false -ErrorAction Stop
        }
        'AllowAll' {
            Set-CsTenantFederationConfiguration -AllowFederatedUsers:$true -AllowedDomains (New-CsEdgeAllowAllKnownDomains) -ErrorAction Stop
        }
        'AllowSpecific' {
            $patterns = $allowedDomains | ForEach-Object { New-CsEdgeDomainPattern -Domain $_ }
            $allowList = New-CsEdgeAllowList -AllowedDomain $patterns
            Set-CsTenantFederationConfiguration -AllowFederatedUsers:$true -AllowedDomains $allowList -ErrorAction Stop
        }
        default {
            throw "Teams-RestrictFederation: unrecognized mode '$mode'. Expected AllowAll, AllowSpecific, or BlockAll."
        }
    }

    return [pscustomobject]@{ Id = 'Teams-RestrictFederation'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = "Updated federation mode to '$mode'." }
}

# ---------------------------------------------------------------------------
# Teams-MeetingJoinDefaults
# ---------------------------------------------------------------------------

function Get-Teams-MeetingJoinDefaultsState {
    <#
    .SYNOPSIS
        Reads the Global meeting policy's lobby bypass and anonymous
        join/start settings.
    .EXAMPLE
        Get-Teams-MeetingJoinDefaultsState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-CsTeamsMeetingPolicy -Identity Global -ErrorAction Stop
    $value = [pscustomobject]@{
        autoAdmittedUsers                = [string]$policy.AutoAdmittedUsers
        allowAnonymousUsersToJoinMeeting  = [bool]$policy.AllowAnonymousUsersToJoinMeeting
        allowAnonymousUsersToStartMeeting = [bool]$policy.AllowAnonymousUsersToStartMeeting
        allowPSTNUsersToBypassLobby       = [bool]$policy.AllowPSTNUsersToBypassLobby
    }
    return [pscustomobject]@{ Id = 'Teams-MeetingJoinDefaults'; Value = $value }
}

function Set-Teams-MeetingJoinDefaultsState {
    <#
    .SYNOPSIS
        Idempotently sets the Global meeting policy's lobby bypass and
        anonymous join/start settings.
    .PARAMETER DesiredValue
        Object: { autoAdmittedUsers: string, allowAnonymousUsersToJoinMeeting: bool,
        allowAnonymousUsersToStartMeeting: bool, allowPSTNUsersToBypassLobby: bool }.
        Both new fields are confirmed-current Boolean parameters of
        Set-CsTeamsMeetingPolicy as of this writing (verified against
        Microsoft Learn's MicrosoftTeams module reference before implementing,
        per the caveat this control shipped with in the v2 gap analysis):
        -AllowAnonymousUsersToStartMeeting controls whether an anonymous
        participant can start (not just join) a meeting; -AllowPSTNUsersToBypassLobby
        controls whether a caller dialing in by phone number bypasses the
        lobby once an authenticated user has joined.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-Teams-MeetingJoinDefaultsState -DesiredValue ([pscustomobject]@{autoAdmittedUsers='EveryoneInCompanyExcludingGuests';allowAnonymousUsersToJoinMeeting=$false;allowAnonymousUsersToStartMeeting=$false;allowPSTNUsersToBypassLobby=$false})
    .NOTES
        Documented MicrosoftTeams module issue: Set-CsTeamsMeetingPolicy can report
        a false-positive HTTP 40301 "Forbidden" when called with -ErrorAction Stop,
        even though the underlying change is applied successfully - the same class
        of bug as the one found in Get-ExternalInOutlook (see
        ExchangeOnlineControls.psm1). See
        https://learn.microsoft.com/answers/questions/5819041. So this call is made
        without an explicit -ErrorAction, and on any error we verify by reading the
        policy back rather than trusting the cmdlet's reported failure.
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
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-Teams-MeetingJoinDefaultsState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'Teams-MeetingJoinDefaults'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    try {
        Set-CsTeamsMeetingPolicy -Identity Global `
            -AutoAdmittedUsers ([string]$DesiredValue.autoAdmittedUsers) `
            -AllowAnonymousUsersToJoinMeeting:([bool]$DesiredValue.allowAnonymousUsersToJoinMeeting) `
            -AllowAnonymousUsersToStartMeeting:([bool]$DesiredValue.allowAnonymousUsersToStartMeeting) `
            -AllowPSTNUsersToBypassLobby:([bool]$DesiredValue.allowPSTNUsersToBypassLobby)
    }
    catch {
        $verify = (Get-Teams-MeetingJoinDefaultsState).Value
        if (-not (Compare-BaselineValueDeep -Left $verify -Right $DesiredValue)) { throw }
    }
    return [pscustomobject]@{ Id = 'Teams-MeetingJoinDefaults'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated CsTeamsMeetingPolicy Global.' }
}

# ---------------------------------------------------------------------------
# Teams-AppPermissionPolicy
# ---------------------------------------------------------------------------

function Get-Teams-AppPermissionPolicyState {
    <#
    .SYNOPSIS
        Reads the Global app permission policy's third-party app catalog behavior.
    .EXAMPLE
        Get-Teams-AppPermissionPolicyState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-CsTeamsAppPermissionPolicy -Identity Global -ErrorAction Stop
    return [pscustomobject]@{ Id = 'Teams-AppPermissionPolicy'; Value = [pscustomobject]@{ globalCatalogAppsType = [string]$policy.GlobalCatalogAppsType } }
}

function Set-Teams-AppPermissionPolicyState {
    <#
    .SYNOPSIS
        Idempotently sets the Global app permission policy's third-party app catalog behavior.
    .PARAMETER DesiredValue
        Object: { globalCatalogAppsType: 'AllowedAppList'|'BlockedAppList'|'AllowAllApps'|'BlockAllApps' }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-Teams-AppPermissionPolicyState -DesiredValue ([pscustomobject]@{globalCatalogAppsType='BlockedAppList'})
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
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-Teams-AppPermissionPolicyState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'Teams-AppPermissionPolicy'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-CsTeamsAppPermissionPolicy -Identity Global -GlobalCatalogAppsType ([string]$DesiredValue.globalCatalogAppsType) -ErrorAction Stop
    return [pscustomobject]@{ Id = 'Teams-AppPermissionPolicy'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated CsTeamsAppPermissionPolicy Global.' }
}

# ---------------------------------------------------------------------------
# Teams-GuestAccessDefault
# ---------------------------------------------------------------------------

function Get-Teams-GuestAccessDefaultState {
    <#
    .SYNOPSIS
        Reads the tenant-wide default guest access toggle.
    .EXAMPLE
        Get-Teams-GuestAccessDefaultState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $cfg = Get-CsTeamsClientConfiguration -Identity Global -ErrorAction Stop
    return [pscustomobject]@{ Id = 'Teams-GuestAccessDefault'; Value = [bool]$cfg.AllowGuestUser }
}

function Set-Teams-GuestAccessDefaultState {
    <#
    .SYNOPSIS
        Idempotently sets the tenant-wide default guest access toggle.
    .PARAMETER DesiredValue
        Boolean.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-Teams-GuestAccessDefaultState -DesiredValue $false
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
    $current = if ($null -ne $CurrentValue) { [bool]$CurrentValue } else { (Get-Teams-GuestAccessDefaultState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'Teams-GuestAccessDefault'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-CsTeamsClientConfiguration -Identity Global -AllowGuestUser:$DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'Teams-GuestAccessDefault'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated CsTeamsClientConfiguration Global.' }
}

Export-ModuleMember -Function @(
    'Get-Teams-BlockConsumerContactState', 'Set-Teams-BlockConsumerContactState'
    'Get-Teams-RestrictFederationState', 'Set-Teams-RestrictFederationState'
    'Get-Teams-MeetingJoinDefaultsState', 'Set-Teams-MeetingJoinDefaultsState'
    'Get-Teams-AppPermissionPolicyState', 'Set-Teams-AppPermissionPolicyState'
    'Get-Teams-GuestAccessDefaultState', 'Set-Teams-GuestAccessDefaultState'
)
