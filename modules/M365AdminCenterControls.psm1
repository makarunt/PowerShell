#Requires -Version 7.0
<#
    M365AdminCenterControls.psm1

    Get-/Set- function pairs for controls that live in the Microsoft 365 admin
    center's org-wide "Org settings" surface rather than under a specific
    workload's own admin center (Entra, Exchange, Teams, SharePoint). This is a
    distinct workload from those four: it's likely to gain more controls over
    time as new org-settings toggles are added to the toolkit, so it gets its
    own dedicated (if currently small) module rather than being folded into one
    of the existing ones.

    None of the controls in this module currently need a live connection - the
    only control here today (M365AdminCenter-SwayExternalSharing) has no
    PowerShell/Graph API and is audit-only. If a future control in this module
    does need a live check, it most likely goes through Microsoft Graph, the
    same way EntraIdControls.psm1 does - do not build a separate connection
    path speculatively; add it only when a real automatable control needs it.
#>

Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# M365AdminCenter-SwayExternalSharing (audit-only; no PowerShell/Graph API)
# ---------------------------------------------------------------------------

function Get-M365AdminCenter-SwayExternalSharingState {
    <#
    .SYNOPSIS
        Audit-only placeholder: whether Sway external sharing is disabled has
        no PowerShell or Microsoft Graph API surface as of this writing - it is
        only configurable through the Microsoft 365 admin center UI. Always
        reports Unknown.
    .EXAMPLE
        Get-M365AdminCenter-SwayExternalSharingState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    return [pscustomobject]@{
        Id     = 'M365AdminCenter-SwayExternalSharing'
        Value  = $null
        Detail = 'No PowerShell or Microsoft Graph API exists for this setting as of this writing; verify manually in the Microsoft 365 admin center.'
    }
}

function Set-M365AdminCenter-SwayExternalSharingState {
    <#
    .SYNOPSIS
        Not automatable: no API exists for this setting. Always returns
        Skipped-Manual.
    .PARAMETER DesiredValue
        Ignored.
    .PARAMETER CurrentValue
        Echoed back for the log/report.
    .EXAMPLE
        Set-M365AdminCenter-SwayExternalSharingState -DesiredValue $false -CurrentValue $null
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
        Id            = 'M365AdminCenter-SwayExternalSharing'
        Status        = 'Skipped-Manual'
        PreviousValue = $CurrentValue
        AppliedValue  = $null
        Message       = 'No API exists for this setting. Change manually: Microsoft 365 admin center > Settings > Org settings > Services > Sway > uncheck "Let people in your organization share their sways with people outside your organization."'
    }
}

Export-ModuleMember -Function @(
    'Get-M365AdminCenter-SwayExternalSharingState', 'Set-M365AdminCenter-SwayExternalSharingState'
)
