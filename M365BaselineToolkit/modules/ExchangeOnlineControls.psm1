#Requires -Version 7.0
<#
    ExchangeOnlineControls.psm1

    Get-/Set- function pairs for every Exchange Online control in the baseline
    inventory. Requires ExchangeOnlineManagement to be connected before use
    (Connect-BaselineWorkload -Connection ExchangeOnline).
#>

Set-StrictMode -Version Latest

function Get-BaselineDefaultAntiPhishPolicyIdentity {
    <#
    .SYNOPSIS
        Resolves the Identity of the tenant's built-in default anti-phish policy.
    .DESCRIPTION
        The built-in policy is conventionally named "Office365 AntiPhish Default",
        not "Default" - Get-AntiPhishPolicy -Identity Default does not resolve it.
        Looks the policy up by its IsDefault flag instead of hardcoding that name,
        in case it's ever renamed, falling back to the documented name if the flag
        isn't present on an older module version.
    .EXAMPLE
        Get-BaselineDefaultAntiPhishPolicyIdentity
    #>
    [CmdletBinding()]
    [OutputType([string])]
    param()
    $default = Get-AntiPhishPolicy -ErrorAction Stop | Where-Object { $_.IsDefault } | Select-Object -First 1
    if ($default) { return $default.Identity }
    return 'Office365 AntiPhish Default'
}

# ---------------------------------------------------------------------------
# ExchangeOnline-MailboxAuditingDefault
# ---------------------------------------------------------------------------

function Get-ExchangeOnline-MailboxAuditingDefaultState {
    <#
    .SYNOPSIS
        Reads whether mailbox auditing is enabled by default for new mailboxes.
    .EXAMPLE
        Get-ExchangeOnline-MailboxAuditingDefaultState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $cfg = Get-OrganizationConfig -ErrorAction Stop
    # AuditDisabled is the raw org-config flag; expose the positive ("enabled") sense
    # to match the desiredValue = true semantics in config.
    return [pscustomobject]@{ Id = 'ExchangeOnline-MailboxAuditingDefault'; Value = -not [bool]$cfg.AuditDisabled }
}

function Set-ExchangeOnline-MailboxAuditingDefaultState {
    <#
    .SYNOPSIS
        Idempotently sets whether mailbox auditing is enabled by default.
    .PARAMETER DesiredValue
        Boolean: true to enable auditing by default.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-ExchangeOnline-MailboxAuditingDefaultState -DesiredValue $true
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
    $current = if ($null -ne $CurrentValue) { [bool]$CurrentValue } else { (Get-ExchangeOnline-MailboxAuditingDefaultState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'ExchangeOnline-MailboxAuditingDefault'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-OrganizationConfig -AuditDisabled:(-not $DesiredValue) -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-MailboxAuditingDefault'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated AuditDisabled.' }
}

# ---------------------------------------------------------------------------
# ExchangeOnline-AntiSpamInbound
# ---------------------------------------------------------------------------

function Get-ExchangeOnline-AntiSpamInboundState {
    <#
    .SYNOPSIS
        Reads the default inbound anti-spam policy's bulk threshold and actions.
    .EXAMPLE
        Get-ExchangeOnline-AntiSpamInboundState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-HostedContentFilterPolicy -Identity Default -ErrorAction Stop
    $value = [pscustomobject]@{
        bulkThreshold            = [int]$policy.BulkThreshold
        highConfidenceSpamAction = [string]$policy.HighConfidenceSpamAction
        spamAction               = [string]$policy.SpamAction
    }
    return [pscustomobject]@{ Id = 'ExchangeOnline-AntiSpamInbound'; Value = $value }
}

function Set-ExchangeOnline-AntiSpamInboundState {
    <#
    .SYNOPSIS
        Idempotently sets the default inbound anti-spam policy's thresholds/actions.
    .PARAMETER DesiredValue
        Object: { bulkThreshold, highConfidenceSpamAction, spamAction }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-ExchangeOnline-AntiSpamInboundState -DesiredValue ([pscustomobject]@{bulkThreshold=6;highConfidenceSpamAction='Quarantine';spamAction='Quarantine'})
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
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-ExchangeOnline-AntiSpamInboundState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'ExchangeOnline-AntiSpamInbound'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-HostedContentFilterPolicy -Identity Default `
        -BulkThreshold ([int]$DesiredValue.bulkThreshold) `
        -HighConfidenceSpamAction ([string]$DesiredValue.highConfidenceSpamAction) `
        -SpamAction ([string]$DesiredValue.spamAction) `
        -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-AntiSpamInbound'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated HostedContentFilterPolicy Default.' }
}

# ---------------------------------------------------------------------------
# ExchangeOnline-AntiPhishing
# ---------------------------------------------------------------------------

function Get-ExchangeOnline-AntiPhishingState {
    <#
    .SYNOPSIS
        Reads the default anti-phishing policy's spoof/mailbox intelligence settings.
    .EXAMPLE
        Get-ExchangeOnline-AntiPhishingState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-AntiPhishPolicy -Identity (Get-BaselineDefaultAntiPhishPolicyIdentity) -ErrorAction Stop
    $value = [pscustomobject]@{
        spoofIntelligence             = [bool]$policy.EnableSpoofIntelligence
        mailboxIntelligence           = [bool]$policy.EnableMailboxIntelligence
        mailboxIntelligenceProtection = [bool]$policy.EnableMailboxIntelligenceProtection
    }
    return [pscustomobject]@{ Id = 'ExchangeOnline-AntiPhishing'; Value = $value }
}

function Set-ExchangeOnline-AntiPhishingState {
    <#
    .SYNOPSIS
        Idempotently sets the default anti-phishing policy's intelligence settings.
    .PARAMETER DesiredValue
        Object: { spoofIntelligence, mailboxIntelligence, mailboxIntelligenceProtection } (all bool).
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-ExchangeOnline-AntiPhishingState -DesiredValue ([pscustomobject]@{spoofIntelligence=$true;mailboxIntelligence=$true;mailboxIntelligenceProtection=$true})
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
    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-ExchangeOnline-AntiPhishingState).Value }
    if (Compare-BaselineValueDeep -Left $current -Right $DesiredValue) {
        return [pscustomobject]@{ Id = 'ExchangeOnline-AntiPhishing'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-AntiPhishPolicy -Identity (Get-BaselineDefaultAntiPhishPolicyIdentity) `
        -EnableSpoofIntelligence:([bool]$DesiredValue.spoofIntelligence) `
        -EnableMailboxIntelligence:([bool]$DesiredValue.mailboxIntelligence) `
        -EnableMailboxIntelligenceProtection:([bool]$DesiredValue.mailboxIntelligenceProtection) `
        -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-AntiPhishing'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated AntiPhishPolicy Default.' }
}

# ---------------------------------------------------------------------------
# ExchangeOnline-AntiMalwareAttachmentFilter
# ---------------------------------------------------------------------------

function Get-ExchangeOnline-AntiMalwareAttachmentFilterState {
    <#
    .SYNOPSIS
        Reads whether the default malware filter policy blocks common dangerous
        attachment file types.
    .EXAMPLE
        Get-ExchangeOnline-AntiMalwareAttachmentFilterState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-MalwareFilterPolicy -Identity Default -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-AntiMalwareAttachmentFilter'; Value = [bool]$policy.EnableFileFilter }
}

function Set-ExchangeOnline-AntiMalwareAttachmentFilterState {
    <#
    .SYNOPSIS
        Idempotently enables/disables the common attachment type filter.
    .PARAMETER DesiredValue
        Boolean.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-ExchangeOnline-AntiMalwareAttachmentFilterState -DesiredValue $true
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
    $current = if ($null -ne $CurrentValue) { [bool]$CurrentValue } else { (Get-ExchangeOnline-AntiMalwareAttachmentFilterState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'ExchangeOnline-AntiMalwareAttachmentFilter'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-MalwareFilterPolicy -Identity Default -EnableFileFilter:$DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-AntiMalwareAttachmentFilter'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated MalwareFilterPolicy Default.' }
}

# ---------------------------------------------------------------------------
# ExchangeOnline-ExternalSenderTag
# ---------------------------------------------------------------------------

function Get-ExchangeOnline-ExternalSenderTagState {
    <#
    .SYNOPSIS
        Reads whether Outlook tags external-sender messages.
    .DESCRIPTION
        Confirmed against a live tenant: Get-ExternalInOutlook's backing endpoint
        (a /adminapi/beta/... REST call, per its own exception stack trace) returns
        HTTP 403 when called immediately after a burst of several other Exchange
        Online reads in the same session - reproduced standalone, entirely outside
        this toolkit, by firing the other ExchangeOnline-* controls' cmdlets first
        and then calling Get-ExternalInOutlook right after; an isolated call with no
        preceding burst always succeeds. This audit calls 7-8 other Exchange Online
        cmdlets in the second or two before it reaches this control, which is enough
        to trip whatever rate limit that beta endpoint has - Microsoft's own error
        handling then fails to decode the real 403 response body (a
        compression/JSON-parsing bug in its own fallback path) and reports a generic
        "server side error" instead of the real reason. A short pause before the
        first attempt gives that window time to clear; -ErrorAction SilentlyContinue
        on the call itself keeps a mid-retry Write-Error inside
        Get-ExternalInOutlook's own implementation from aborting under this
        toolkit's global $ErrorActionPreference = 'Stop'.
    .EXAMPLE
        Get-ExchangeOnline-ExternalSenderTagState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    Start-Sleep -Seconds 5
    $maxAttempts = 3
    $lastError = $null
    for ($attempt = 1; $attempt -le $maxAttempts; $attempt++) {
        $cfg = Get-ExternalInOutlook -ErrorAction SilentlyContinue -ErrorVariable getError
        if ($cfg) {
            return [pscustomobject]@{ Id = 'ExchangeOnline-ExternalSenderTag'; Value = [bool]$cfg.Enabled }
        }
        # Plain string conversion rather than .Exception.Message: whatever lands in
        # $getError[0] (a normal ErrorRecord most of the time) always has a sane
        # ToString(), so this can't itself throw under StrictMode the way a property
        # chain that assumes one specific object shape can.
        $lastError = if ($getError -and $getError.Count -gt 0) { [string]$getError[0] } else { 'no result returned' }
        if ($attempt -lt $maxAttempts) { Start-Sleep -Seconds 10 }
    }
    throw "Get-ExternalInOutlook failed after $maxAttempts attempt(s), each preceded by a pause to clear any rate limit from preceding calls: $lastError"
}

function Set-ExchangeOnline-ExternalSenderTagState {
    <#
    .SYNOPSIS
        Idempotently enables/disables the external sender tag feature. Falls back
        to New-ExternalInOutlook if the feature has never been initialized in this
        tenant (Set-ExternalInOutlook requires an existing configuration object).
    .PARAMETER DesiredValue
        Boolean.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-ExchangeOnline-ExternalSenderTagState -DesiredValue $true
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
    $current = if ($null -ne $CurrentValue) { [bool]$CurrentValue } else { (Get-ExchangeOnline-ExternalSenderTagState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'ExchangeOnline-ExternalSenderTag'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    try {
        Set-ExternalInOutlook -Enabled:$DesiredValue -ErrorAction Stop
    }
    catch {
        New-ExternalInOutlook -Enabled:$DesiredValue -ErrorAction Stop
    }
    return [pscustomobject]@{ Id = 'ExchangeOnline-ExternalSenderTag'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated ExternalInOutlook.' }
}

# ---------------------------------------------------------------------------
# ExchangeOnline-DisableAutoForwarding
# ---------------------------------------------------------------------------

function Get-ExchangeOnline-DisableAutoForwardingState {
    <#
    .SYNOPSIS
        Reads the tenant-wide outbound auto-forwarding mode.
    .EXAMPLE
        Get-ExchangeOnline-DisableAutoForwardingState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $policy = Get-HostedOutboundSpamFilterPolicy -Identity Default -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-DisableAutoForwarding'; Value = [string]$policy.AutoForwardingMode }
}

function Set-ExchangeOnline-DisableAutoForwardingState {
    <#
    .SYNOPSIS
        Idempotently sets the tenant-wide outbound auto-forwarding mode.
    .PARAMETER DesiredValue
        String enum, e.g. 'Off'.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-ExchangeOnline-DisableAutoForwardingState -DesiredValue 'Off'
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
    $current = if ($null -ne $CurrentValue) { [string]$CurrentValue } else { (Get-ExchangeOnline-DisableAutoForwardingState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'ExchangeOnline-DisableAutoForwarding'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-HostedOutboundSpamFilterPolicy -Identity Default -AutoForwardingMode $DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-DisableAutoForwarding'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated AutoForwardingMode.' }
}

# ---------------------------------------------------------------------------
# ExchangeOnline-DisableSmtpAuth
# ---------------------------------------------------------------------------

function Get-ExchangeOnline-DisableSmtpAuthState {
    <#
    .SYNOPSIS
        Reads whether tenant-wide SMTP AUTH client authentication is disabled.
    .EXAMPLE
        Get-ExchangeOnline-DisableSmtpAuthState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $cfg = Get-TransportConfig -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-DisableSmtpAuth'; Value = [bool]$cfg.SmtpClientAuthenticationDisabled }
}

function Set-ExchangeOnline-DisableSmtpAuthState {
    <#
    .SYNOPSIS
        Idempotently disables/enables tenant-wide SMTP AUTH.
    .PARAMETER DesiredValue
        Boolean: true to disable SMTP AUTH tenant-wide.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-ExchangeOnline-DisableSmtpAuthState -DesiredValue $true
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
    $current = if ($null -ne $CurrentValue) { [bool]$CurrentValue } else { (Get-ExchangeOnline-DisableSmtpAuthState).Value }
    if ($current -eq $DesiredValue) {
        return [pscustomobject]@{ Id = 'ExchangeOnline-DisableSmtpAuth'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    Set-TransportConfig -SmtpClientAuthenticationDisabled:$DesiredValue -ErrorAction Stop
    return [pscustomobject]@{ Id = 'ExchangeOnline-DisableSmtpAuth'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = 'Updated SmtpClientAuthenticationDisabled.' }
}

# ---------------------------------------------------------------------------
# ExchangeOnline-DkimSigning
# ---------------------------------------------------------------------------

function Get-ExchangeOnline-DkimSigningState {
    <#
    .SYNOPSIS
        Reads DKIM signing state for every configured accepted domain.
    .EXAMPLE
        Get-ExchangeOnline-DkimSigningState
    #>
    [CmdletBinding()]
    [OutputType([pscustomobject])]
    param()
    $configs = @(Get-DkimSigningConfig -ErrorAction Stop)
    $domains = @($configs | ForEach-Object { [pscustomobject]@{ domain = [string]$_.Domain; enabled = [bool]$_.Enabled } })
    return [pscustomobject]@{ Id = 'ExchangeOnline-DkimSigning'; Value = [pscustomobject]@{ domains = $domains } }
}

function Set-ExchangeOnline-DkimSigningState {
    <#
    .SYNOPSIS
        Idempotently enables DKIM signing for every domain named in DesiredValue.domains.
        Creates a new signing config for a domain that has none, or enables an
        existing one that is currently disabled. Requires a non-empty domain list -
        this control is tenant-specific and has no safe universal default; the
        orchestrator's pre-flight check (Test-BaselineApplyReadiness) should already
        have blocked an Apply run with an empty list, but this function guards
        defensively too.
    .PARAMETER DesiredValue
        Object: { domains: [ 'contoso.com', ... ] }.
    .PARAMETER CurrentValue
        Optional pre-fetched current value.
    .EXAMPLE
        Set-ExchangeOnline-DkimSigningState -DesiredValue ([pscustomobject]@{domains=@('contoso.com')})
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

    $domainList = @($DesiredValue.domains)
    if ($domainList.Count -eq 0) {
        throw "ExchangeOnline-DkimSigning requires at least one domain in desiredValue.domains; update config/baseline.config.json before running Apply."
    }

    $current = if ($null -ne $CurrentValue) { $CurrentValue } else { (Get-ExchangeOnline-DkimSigningState).Value }

    $changed = $false
    foreach ($domain in $domainList) {
        $existing = Get-DkimSigningConfig -Identity $domain -ErrorAction SilentlyContinue
        if ($existing) {
            if (-not $existing.Enabled) {
                Set-DkimSigningConfig -Identity $domain -Enabled $true -ErrorAction Stop
                $changed = $true
            }
        }
        else {
            New-DkimSigningConfig -DomainName $domain -Enabled $true -ErrorAction Stop
            $changed = $true
        }
    }

    if (-not $changed) {
        return [pscustomobject]@{ Id = 'ExchangeOnline-DkimSigning'; Status = 'Success'; PreviousValue = $current; AppliedValue = $current; Message = 'Already compliant (no-op).' }
    }
    return [pscustomobject]@{ Id = 'ExchangeOnline-DkimSigning'; Status = 'Success'; PreviousValue = $current; AppliedValue = $DesiredValue; Message = "Enabled DKIM signing for: $($domainList -join ', ')." }
}

Export-ModuleMember -Function @(
    'Get-ExchangeOnline-MailboxAuditingDefaultState', 'Set-ExchangeOnline-MailboxAuditingDefaultState'
    'Get-ExchangeOnline-AntiSpamInboundState', 'Set-ExchangeOnline-AntiSpamInboundState'
    'Get-ExchangeOnline-AntiPhishingState', 'Set-ExchangeOnline-AntiPhishingState'
    'Get-ExchangeOnline-AntiMalwareAttachmentFilterState', 'Set-ExchangeOnline-AntiMalwareAttachmentFilterState'
    'Get-ExchangeOnline-ExternalSenderTagState', 'Set-ExchangeOnline-ExternalSenderTagState'
    'Get-ExchangeOnline-DisableAutoForwardingState', 'Set-ExchangeOnline-DisableAutoForwardingState'
    'Get-ExchangeOnline-DisableSmtpAuthState', 'Set-ExchangeOnline-DisableSmtpAuthState'
    'Get-ExchangeOnline-DkimSigningState', 'Set-ExchangeOnline-DkimSigningState'
)
