#Requires -Version 7.0
<#
    SharePointOnlineControls.psm1

    Get-/Set- function pairs for every SharePoint Online (and OneDrive, which
    shares the same tenant-level settings) control in the baseline inventory.
    Requires Microsoft.Online.SharePoint.PowerShell to be connected before use
    (Connect-BaselineWorkload -Connection SharePointOnline -SharePointAdminUrl ...).
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# SharePointOnline-SharingCapability
# ---------------------------------------------------------------------------

function Get-SharePointOnline-SharingCapabilityState {
    <#
    .SYNOPSIS
        Reads the tenant-wide external sharing ceiling.
    .EXAMPLE
        Get-SharePointOnline-SharingCapabilityState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $tenant = Get-SPOTenant -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-SharingCapability'; Value = [string]$tenant.SharingCapability }
}

function Set-SharePointOnline-SharingCapabilityState {
    <#
    .SYNOPSIS
        Idempotently sets the tenant-wide external sharing ceiling.
    .PARAMETER DesiredValue
        String enum accepted by Set-SPOTenant -SharingCapability.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-SharePointOnline-SharingCapabilityState -DesiredValue 'ExternalUserSharingOnly'
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
    $current = if ($null -ne $CurrentValue) { [string]$CurrentValue } else { (Get-SharePointOnline-SharingCapabilityState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'SharePointOnline-SharingCapability'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-SPOTenant -SharingCapability $DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-SharingCapability'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated SharingCapability.' }
}

# ---------------------------------------------------------------------------
# SharePointOnline-DefaultSharingLinkType
# ---------------------------------------------------------------------------

function Get-SharePointOnline-DefaultSharingLinkTypeState {
    <#
    .SYNOPSIS
        Reads the default sharing link type.
    .EXAMPLE
        Get-SharePointOnline-DefaultSharingLinkTypeState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $tenant = Get-SPOTenant -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-DefaultSharingLinkType'; Value = [string]$tenant.DefaultSharingLinkType }
}

function Set-SharePointOnline-DefaultSharingLinkTypeState {
    <#
    .SYNOPSIS
        Idempotently sets the default sharing link type.
    .PARAMETER DesiredValue
        String enum accepted by Set-SPOTenant -DefaultSharingLinkType.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-SharePointOnline-DefaultSharingLinkTypeState -DesiredValue 'Direct'
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
    $current = if ($null -ne $CurrentValue) { [string]$CurrentValue } else { (Get-SharePointOnline-DefaultSharingLinkTypeState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'SharePointOnline-DefaultSharingLinkType'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-SPOTenant -DefaultSharingLinkType $DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-DefaultSharingLinkType'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated DefaultSharingLinkType.' }
}

# ---------------------------------------------------------------------------
# SharePointOnline-DefaultLinkPermission
# ---------------------------------------------------------------------------

function Get-SharePointOnline-DefaultLinkPermissionState {
    <#
    .SYNOPSIS
        Reads the default permission granted by a new sharing link.
    .EXAMPLE
        Get-SharePointOnline-DefaultLinkPermissionState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $tenant = Get-SPOTenant -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-DefaultLinkPermission'; Value = [string]$tenant.DefaultLinkPermission }
}

function Set-SharePointOnline-DefaultLinkPermissionState {
    <#
    .SYNOPSIS
        Idempotently sets the default permission granted by a new sharing link.
    .PARAMETER DesiredValue
        String enum accepted by Set-SPOTenant -DefaultLinkPermission.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-SharePointOnline-DefaultLinkPermissionState -DesiredValue 'View'
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
    $current = if ($null -ne $CurrentValue) { [string]$CurrentValue } else { (Get-SharePointOnline-DefaultLinkPermissionState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'SharePointOnline-DefaultLinkPermission'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    # 'None' is a value Get-SPOTenant can report (SharePoint's own never-configured
    # default - e.g. what a Restore snapshot captured before this control was ever
    # applied) but Set-SPOTenant -DefaultLinkPermission rejects as input; it only
    # accepts 'View' or 'Edit'. There's no supported way to programmatically revert
    # to 'None' once a value has been set, so fail with a clear reason up front
    # instead of letting Set-SPOTenant's own opaque validation error surface.
    if ($DesiredValue -notin @('View', 'Edit')) {
        $message = "Cannot set DefaultLinkPermission to '$DesiredValue' - Set-SPOTenant only accepts 'View' or 'Edit'. '$DesiredValue' is SharePoint's own unconfigured default and cannot be restored via PowerShell; change it manually in the SharePoint admin center if you need this reverted."
        return [pscustomobject]@{ Id = 'SharePointOnline-DefaultLinkPermission'; Status = 'Failed'; PreviousValue = $current; AppliedValue = $null; Message = $message }
    }
    Set-SPOTenant -DefaultLinkPermission $DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-DefaultLinkPermission'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated DefaultLinkPermission.' }
}

# ---------------------------------------------------------------------------
# SharePointOnline-AnonymousLinkExpiration
# ---------------------------------------------------------------------------

function Get-SharePointOnline-AnonymousLinkExpirationState {
    <#
    .SYNOPSIS
        Reads the number of days after which anonymous 'Anyone' links expire.
    .EXAMPLE
        Get-SharePointOnline-AnonymousLinkExpirationState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $tenant = Get-SPOTenant -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-AnonymousLinkExpiration'; Value = [int]$tenant.RequireAnonymousLinksExpireInDays }
}

function Set-SharePointOnline-AnonymousLinkExpirationState {
    <#
    .SYNOPSIS
        Idempotently sets the anonymous link expiration window, in days.
    .PARAMETER DesiredValue
        Integer number of days.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-SharePointOnline-AnonymousLinkExpirationState -DesiredValue 30
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory)]
        [int]$DesiredValue,

        [Parameter()]
        [AllowNull()]
        [object]$CurrentValue
    )
    $current = if ($null -ne $CurrentValue) { [int]$CurrentValue } else { (Get-SharePointOnline-AnonymousLinkExpirationState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'SharePointOnline-AnonymousLinkExpiration'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    # -1 is a value Get-SPOTenant can report (SharePoint's own never-configured
    # default, meaning anonymous links never expire - e.g. what a Restore snapshot
    # captured before this control was ever applied) but Set-SPOTenant rejects as
    # input; it only accepts 1-730. There's no supported way to programmatically
    # revert to -1 once a value has been set, so fail with a clear reason up front
    # instead of letting Set-SPOTenant's own opaque validation error surface.
    if ($DesiredValue -lt 1 -or $DesiredValue -gt 730) {
        $message = "Cannot set RequireAnonymousLinksExpireInDays to $DesiredValue - Set-SPOTenant only accepts values from 1 to 730. $DesiredValue is SharePoint's own unconfigured default (never expire) and cannot be restored via PowerShell; change it manually in the SharePoint admin center if you need this reverted."
        return [pscustomobject]@{ Id = 'SharePointOnline-AnonymousLinkExpiration'; Status = 'Failed'; PreviousValue = $current; AppliedValue = $null; Message = $message }
    }
    Set-SPOTenant -RequireAnonymousLinksExpireInDays $DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-AnonymousLinkExpiration'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated RequireAnonymousLinksExpireInDays.' }
}

# ---------------------------------------------------------------------------
# SharePointOnline-LegacyAuthProtocols
# ---------------------------------------------------------------------------

function Get-SharePointOnline-LegacyAuthProtocolsState {
    <#
    .SYNOPSIS
        Reads whether legacy (non-modern-auth) client protocols are allowed.
    .EXAMPLE
        Get-SharePointOnline-LegacyAuthProtocolsState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $tenant = Get-SPOTenant -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-LegacyAuthProtocols'; Value = [bool]$tenant.LegacyAuthProtocolsEnabled }
}

function Set-SharePointOnline-LegacyAuthProtocolsState {
    <#
    .SYNOPSIS
        Idempotently blocks/allows legacy (non-modern-auth) client protocols.
    .PARAMETER DesiredValue
        Boolean: false to block legacy protocols.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-SharePointOnline-LegacyAuthProtocolsState -DesiredValue $false
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
    $current = if ($null -ne $CurrentValue) { [bool]$CurrentValue } else { (Get-SharePointOnline-LegacyAuthProtocolsState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'SharePointOnline-LegacyAuthProtocols'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-SPOTenant -LegacyAuthProtocolsEnabled:$DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-LegacyAuthProtocols'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated LegacyAuthProtocolsEnabled.' }
}

# ---------------------------------------------------------------------------
# SharePointOnline-IdleSessionSignOut
# ---------------------------------------------------------------------------

function Get-SharePointOnline-IdleSessionSignOutState {
    <#
    .SYNOPSIS
        Reads the browser idle session sign-out policy.
    .EXAMPLE
        Get-SharePointOnline-IdleSessionSignOutState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $cfg = Get-SPOBrowserIdleSignOut -ErrorAction Stop
    $value = [pscustomobject]@{
        enabled              = [bool]$cfg.Enabled
        warnAfterMinutes     = [int]([timespan]$cfg.WarnAfter).TotalMinutes
        signOutAfterMinutes  = [int]([timespan]$cfg.SignOutAfter).TotalMinutes
    }
    return [pscustomobject]@{ Id = 'SharePointOnline-IdleSessionSignOut'; Value = $value }
}

function Set-SharePointOnline-IdleSessionSignOutState {
    <#
    .SYNOPSIS
        Idempotently sets the browser idle session sign-out policy.
    .PARAMETER DesiredValue
        Object: { enabled: bool, warnAfterMinutes: int, signOutAfterMinutes: int }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-SharePointOnline-IdleSessionSignOutState -DesiredValue ([pscustomobject]@{enabled=$true;warnAfterMinutes=15;signOutAfterMinutes=20})
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
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-SharePointOnline-IdleSessionSignOutState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'SharePointOnline-IdleSessionSignOut'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-SPOBrowserIdleSignOut -Enabled:([bool]$DesiredValue.enabled) `
        -WarnAfter (New-TimeSpan -Minutes ([int]$DesiredValue.warnAfterMinutes)) `
        -SignOutAfter (New-TimeSpan -Minutes ([int]$DesiredValue.signOutAfterMinutes)) `
        -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-IdleSessionSignOut'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated SPOBrowserIdleSignOut.' }
}

# ---------------------------------------------------------------------------
# SharePointOnline-UnmanagedDeviceAccess
# ---------------------------------------------------------------------------

function Get-SharePointOnline-UnmanagedDeviceAccessState {
    <#
    .SYNOPSIS
        Reads the unmanaged/unknown device access posture.
    .EXAMPLE
        Get-SharePointOnline-UnmanagedDeviceAccessState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $tenant = Get-SPOTenant -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-UnmanagedDeviceAccess'; Value = [string]$tenant.ConditionalAccessPolicy }
}

function Set-SharePointOnline-UnmanagedDeviceAccessState {
    <#
    .SYNOPSIS
        Idempotently sets the unmanaged/unknown device access posture.
    .PARAMETER DesiredValue
        String enum accepted by Set-SPOTenant -ConditionalAccessPolicy.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-SharePointOnline-UnmanagedDeviceAccessState -DesiredValue 'AllowLimitedAccess'
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
    $current = if ($null -ne $CurrentValue) { [string]$CurrentValue } else { (Get-SharePointOnline-UnmanagedDeviceAccessState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'SharePointOnline-UnmanagedDeviceAccess'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-SPOTenant -ConditionalAccessPolicy $DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'SharePointOnline-UnmanagedDeviceAccess'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated ConditionalAccessPolicy.' }
}

Export-ModuleMember -Function @(
    'Get-SharePointOnline-SharingCapabilityState', 'Set-SharePointOnline-SharingCapabilityState'
    'Get-SharePointOnline-DefaultSharingLinkTypeState', 'Set-SharePointOnline-DefaultSharingLinkTypeState'
    'Get-SharePointOnline-DefaultLinkPermissionState', 'Set-SharePointOnline-DefaultLinkPermissionState'
    'Get-SharePointOnline-AnonymousLinkExpirationState', 'Set-SharePointOnline-AnonymousLinkExpirationState'
    'Get-SharePointOnline-LegacyAuthProtocolsState', 'Set-SharePointOnline-LegacyAuthProtocolsState'
    'Get-SharePointOnline-IdleSessionSignOutState', 'Set-SharePointOnline-IdleSessionSignOutState'
    'Get-SharePointOnline-UnmanagedDeviceAccessState', 'Set-SharePointOnline-UnmanagedDeviceAccessState'
)
