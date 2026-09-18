#Requires -Version 7.0
<#
.SYNOPSIS
    Applies, audits, or rolls back the M365 Baseline Toolkit's minimum-viable
    security/governance baseline, connecting with app-only (certificate)
    authentication instead of interactive sign-in.
.DESCRIPTION
    Same config file, same control catalog (all five workload modules), same
    reports/backups folders and file-naming convention as
    Invoke-M365Baseline.ps1 - the only difference is how this script connects
    to Graph, Exchange Online, Teams, and SharePoint Online. It never touches
    Windows Account Manager (WAM), never opens a browser, and has no
    interactive fallback anywhere - if app-only auth doesn't work for a given
    service or a given control's cmdlet, this script fails loudly rather than
    silently falling back to an interactive prompt.

    Architecture note: BaselineCore.psm1 already separates "establish
    connections" (Connect-BaselineWorkload, called only from
    Invoke-M365Baseline.ps1's top-level script body) from "run the
    Audit/Apply/Restore loop" (Invoke-BaselineControlAudit/-Apply/-Restore,
    which take an already-connected session as a given and never call any
    Connect-* cmdlet themselves). Because of that, this script needed no copy
    of BaselineCore.psm1's orchestration engine - it imports and calls the
    exact same exported functions the interactive script does
    (Import-BaselineConfig, Get-BaselineControlCatalog,
    Get-BaselineConnectionOrder, Assert-BaselineRequiredModules,
    Invoke-BaselineControlAudit/-Apply/-Restore, Save-/Import-BaselineSnapshot,
    Export-BaselineMarkdownReport/-HtmlReport, Disconnect-BaselineWorkload) -
    only the connection step itself is swapped, via
    modules/AppOnlyConnections.psm1's Connect-M365BaselineServicesAppOnly.
    See README.md's "App-only (certificate) authentication" section for the
    fuller explanation and the one-time tenant setup this requires.
.PARAMETER Mode
    'Audit', 'Apply', or 'Restore'. Mutually exclusive run modes - identical
    contract to Invoke-M365Baseline.ps1.
.PARAMETER AppId
    The Entra app registration's Application (client) ID.
.PARAMETER TenantId
    Entra tenant id (GUID) or a verified domain (e.g. contoso.onmicrosoft.com).
.PARAMETER Organization
    The tenant's primary *.onmicrosoft.com domain. Required only when an
    enabled control needs an ExchangeOnline connection (Exchange Online's
    app-only auth requires this specific verified domain, not a GUID).
.PARAMETER SpoAdminUrl
    SharePoint admin center URL (e.g. https://contoso-admin.sharepoint.com).
    Required only when an enabled control needs a SharePointOnline connection.
.PARAMETER CertificateThumbprint
    Thumbprint of a certificate already imported into a local certificate
    store, whose private key matches the public certificate uploaded to the
    app registration. Mutually exclusive with -CertificatePath/-Certificate.
.PARAMETER CertificateStoreLocation
    'CurrentUser' (default) or 'LocalMachine' - which store
    -CertificateThumbprint is looked up in.
.PARAMETER CertificatePath
    Path to a .pfx file containing the certificate and private key, as an
    alternative to the certificate store (e.g. downloaded from Key Vault to a
    temp path at runtime). Mutually exclusive with
    -CertificateThumbprint/-Certificate.
.PARAMETER CertificatePassword
    SecureString password for -CertificatePath. Omit if the .pfx has none.
.PARAMETER Certificate
    An already-constructed X509Certificate2 object, for a caller that manages
    its own certificate retrieval (Key Vault SDK, etc.). Mutually exclusive
    with -CertificateThumbprint/-CertificatePath.
.PARAMETER ConfigPath
    Path to the desired-state config JSON file. Same default and same file as
    Invoke-M365Baseline.ps1.
.PARAMETER SchemaPath
    Path to the config's JSON schema file.
.PARAMETER ReportPath
    Directory reports are written to (created if missing). Same default
    folder as Invoke-M365Baseline.ps1 - reports/backups from either script are
    interchangeable.
.PARAMETER BackupPath
    Directory backup/snapshot files are written to (created if missing).
.PARAMETER BackupFile
    Required for -Mode Restore: path to a snapshot JSON file produced by a
    prior Audit run or Apply run's pre-change phase, from *either* this script
    or Invoke-M365Baseline.ps1 - both produce and consume the same format.
.PARAMETER InstallMissingModules
    Automatically install any missing required module from PSGallery
    (-Scope CurrentUser) instead of failing with an actionable error.
.PARAMETER StopOnError
    Abort the whole Apply/Restore run on the first control failure instead of
    continuing past it and reporting all failures at the end.
.PARAMETER AcknowledgeFederationBlockAll
    Required to apply Teams-RestrictFederation when its desiredValue.allowedDomains
    is empty with mode 'AllowSpecific' - an intentional but severe "block all
    external federation" configuration.
.PARAMETER IncludeHtmlReport
    Also write an HTML copy of every Markdown report.
.PARAMETER KeepConnectionsOpen
    Skip disconnecting from Graph/Exchange Online/Teams/SharePoint at the end
    of the run - same behavior and same session-global tracking as
    Invoke-M365Baseline.ps1's identical switch.
.EXAMPLE
    ./Invoke-M365Baseline.AppOnly.ps1 -Mode Audit -AppId $appId -TenantId contoso.onmicrosoft.com -CertificateThumbprint $thumbprint
.EXAMPLE
    ./Invoke-M365Baseline.AppOnly.ps1 -Mode Apply -AppId $appId -TenantId contoso.onmicrosoft.com -Organization contoso.onmicrosoft.com -SpoAdminUrl https://contoso-admin.sharepoint.com -CertificatePath ./app-only.pfx -CertificatePassword (Read-Host -AsSecureString) -WhatIf
.NOTES
    Exit codes: same as Invoke-M365Baseline.ps1 - 0 success, 1 startup/
    validation/connection failure, 2 Audit completed with read errors,
    3 Apply/Restore completed with control failures.

    Per-control failures during Apply/Restore are relabeled
    'Failed-AppOnlyUnsupported' in this script's own console output and
    reports (the underlying change log, written by BaselineCore.psm1 itself,
    keeps its original 'Failed' status - that file is not touched by this
    relabeling). This script never retries a failed control interactively;
    if a control's cmdlet turns out not to work under app-only auth for your
    tenant, re-run just that control via Invoke-M365Baseline.ps1 instead.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Thumbprint')]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Audit', 'Apply', 'Restore')]
    [string]$Mode,

    [Parameter(Mandatory)]
    [string]$AppId,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter()]
    [string]$Organization,

    [Parameter()]
    [string]$SpoAdminUrl,

    [Parameter(Mandatory, ParameterSetName = 'Thumbprint')]
    [string]$CertificateThumbprint,

    [Parameter(ParameterSetName = 'Thumbprint')]
    [ValidateSet('CurrentUser', 'LocalMachine')]
    [string]$CertificateStoreLocation = 'CurrentUser',

    [Parameter(Mandatory, ParameterSetName = 'File')]
    [string]$CertificatePath,

    [Parameter(ParameterSetName = 'File')]
    [securestring]$CertificatePassword,

    [Parameter(Mandatory, ParameterSetName = 'Object')]
    [System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,

    [Parameter()]
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config/baseline.config.json'),

    [Parameter()]
    [string]$SchemaPath = (Join-Path $PSScriptRoot 'config/baseline.config.schema.json'),

    [Parameter()]
    [string]$ReportPath = (Join-Path $PSScriptRoot 'reports'),

    [Parameter()]
    [string]$BackupPath = (Join-Path $PSScriptRoot 'backups'),

    [Parameter()]
    [string]$BackupFile,

    [Parameter()]
    [switch]$InstallMissingModules,

    [Parameter()]
    [switch]$StopOnError,

    [Parameter()]
    [switch]$AcknowledgeFederationBlockAll,

    [Parameter()]
    [switch]$IncludeHtmlReport,

    [Parameter()]
    [switch]$KeepConnectionsOpen
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$modulesDir = Join-Path $PSScriptRoot 'modules'
Import-Module (Join-Path $modulesDir 'BaselineCore.psm1') -Force -Global
Import-Module (Join-Path $modulesDir 'AppOnlyConnections.psm1') -Force -Global
Import-Module (Join-Path $modulesDir 'EntraIdControls.psm1') -Force -Global -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesDir 'ExchangeOnlineControls.psm1') -Force -Global -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesDir 'TeamsControls.psm1') -Force -Global -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesDir 'SharePointOnlineControls.psm1') -Force -Global -WarningAction SilentlyContinue
Import-Module (Join-Path $modulesDir 'ConditionalAccessControls.psm1') -Force -Global -WarningAction SilentlyContinue

$exitCode = 0
$connectedServices = [System.Collections.Generic.List[string]]::new()

function Write-BaselineHost {
    param([string]$Message, [string]$Color = 'Gray')
    Write-Host $Message -ForegroundColor $Color
}

function ConvertTo-AppOnlyResultSet {
    <#
    .SYNOPSIS
        Local to this script only: relabels 'Failed' results from
        Invoke-BaselineControlApply/-Restore as 'Failed-AppOnlyUnsupported'
        for this script's own console output and reports, per this script's
        "no interactive fallback" contract - a post-connection control
        failure under app-only auth is actionable differently (re-run that
        control via the interactive script) than a generic failure, and this
        makes that visible without needing BaselineCore.psm1 itself to know
        anything about app-only auth. Does not touch the change log file,
        which BaselineCore.psm1 already wrote with the original status by the
        time this runs.
    #>
    param([Parameter(Mandatory)][AllowEmptyCollection()][pscustomobject[]]$Results)
    return @($Results | ForEach-Object {
        if ($_.Status -eq 'Failed') {
            [pscustomobject]@{
                Id            = $_.Id
                Status        = 'Failed-AppOnlyUnsupported'
                PreviousValue = $_.PreviousValue
                AppliedValue  = $_.AppliedValue
                Message       = "$($_.Message) [app-only run: this script never falls back to interactive auth automatically. If this is a cmdlet/parameter combination that doesn't support app-only auth for your tenant, re-run just this control via the interactive Invoke-M365Baseline.ps1 instead.]"
            }
        }
        else {
            $_
        }
    })
}

try {
    if ($Mode -eq 'Restore' -and [string]::IsNullOrWhiteSpace($BackupFile)) {
        throw "-Mode Restore requires -BackupFile <path to a snapshot JSON file produced by Audit or Apply - from this script or Invoke-M365Baseline.ps1, either is fine>."
    }

    Write-BaselineHost "Loading and validating config: $ConfigPath" 'Cyan'
    $config = Import-BaselineConfig -Path $ConfigPath -SchemaPath $SchemaPath
    Write-BaselineHost "Config valid. $($config.controls.Count) control(s) defined, $(@($config.controls | Where-Object enabled).Count) enabled." 'Green'

    $catalog = Get-BaselineControlCatalog -Config $config
    Write-BaselineHost "Control catalog resolved: $($catalog.Count) enabled control(s) matched to Get-/Set- functions." 'Green'

    if ($Mode -eq 'Apply') {
        $readinessErrors = Test-BaselineApplyReadiness -Catalog $catalog
        if ($readinessErrors.Count -gt 0) {
            $message = "Apply readiness check failed with $($readinessErrors.Count) issue(s) - fix these before Apply can run:`n" + (($readinessErrors | ForEach-Object { " - $_" }) -join "`n")
            throw $message
        }
    }

    # Same union-of-primary-and-extra-connections logic as Invoke-M365Baseline.ps1
    # (reusing Get-BaselineControlCatalog's .Connection/.ExtraConnections fields
    # unchanged) - a Graph connection is established whenever ANY enabled control
    # needs one for ANY reason (e.g. ExchangeOnline-AntiPhishingMailboxIntelligence's
    # license check), not just controls whose primary connection is Graph.
    $neededConnections = @($catalog | ForEach-Object { @($_.Connection) + @($_.ExtraConnections) })
    $requiredConnections = Get-BaselineConnectionOrder -Connections $neededConnections

    Write-BaselineHost "Checking required modules for: $($requiredConnections -join ', ')" 'Cyan'
    Assert-BaselineRequiredModules -Connections $requiredConnections -InstallMissingModules:$InstallMissingModules

    if (($requiredConnections -contains 'ExchangeOnline') -and [string]::IsNullOrWhiteSpace($Organization)) {
        throw "An enabled control requires an ExchangeOnline connection. Pass -Organization <tenant>.onmicrosoft.com."
    }
    if (($requiredConnections -contains 'SharePointOnline') -and [string]::IsNullOrWhiteSpace($SpoAdminUrl)) {
        throw "An enabled control requires a SharePointOnline connection. Pass -SpoAdminUrl https://<tenant>-admin.sharepoint.com."
    }

    # Connect lazily: only to services this run's enabled controls actually need
    # (already narrowed by $requiredConnections above), and only the ones not
    # already left open by -KeepConnectionsOpen on a prior run in this window.
    $toConnect = [System.Collections.Generic.List[string]]::new()
    foreach ($conn in $requiredConnections) {
        if (Test-BaselineWorkloadConnected -Connection $conn) {
            Write-BaselineHost "Reusing existing $conn connection (left open by -KeepConnectionsOpen on a prior run in this window)." 'DarkGray'
        }
        else {
            $toConnect.Add($conn)
        }
        # Tracked either way, same rationale as Invoke-M365Baseline.ps1: a
        # connection reused from an earlier run should still be disconnected at
        # the end of this run too, unless -KeepConnectionsOpen is set here.
        $connectedServices.Add($conn)
    }

    if ($toConnect.Count -gt 0) {
        Write-BaselineHost "Connecting (app-only) to: $($toConnect -join ', ')..." 'Cyan'
        $connectParams = @{ Services = $toConnect.ToArray(); AppId = $AppId; TenantId = $TenantId }
        if ($Organization) { $connectParams['Organization'] = $Organization }
        if ($SpoAdminUrl) { $connectParams['SpoAdminUrl'] = $SpoAdminUrl }
        switch ($PSCmdlet.ParameterSetName) {
            'Thumbprint' {
                $connectParams['CertificateThumbprint'] = $CertificateThumbprint
                $connectParams['CertificateStoreLocation'] = $CertificateStoreLocation
            }
            'File' {
                $connectParams['CertificatePath'] = $CertificatePath
                if ($CertificatePassword) { $connectParams['CertificatePassword'] = $CertificatePassword }
            }
            'Object' {
                $connectParams['Certificate'] = $Certificate
            }
        }
        Connect-M365BaselineServicesAppOnly @connectParams
        foreach ($c in $toConnect) { Write-BaselineHost "Connected to $c (app-only)." 'Green' }
    }

    if ($catalog | Where-Object { $_.Workload -eq 'ConditionalAccess' }) {
        $caSummary = Get-BaselineConditionalAccessSummary
        Write-BaselineHost "`nConditional Access controls:" 'Yellow'
        if (-not $caSummary.Tier1Available) {
            Write-BaselineHost "  Skipped: tenant has Entra ID Free, which does not support Conditional Access. No CA policy will be read or changed this run." 'Yellow'
        }
        else {
            Write-BaselineHost "  Tier 1 (Entra ID P1) controls: evaluated." 'Yellow'
            Write-BaselineHost "  Tier 2 (Entra ID P2) controls: $(if ($caSummary.Tier2Available) { 'evaluated.' } else { 'Skipped-LicenseInsufficient - tenant does not have Entra ID P2.' })" 'Yellow'
            $groupState = if ($caSummary.EmergencyGroupId) { "exists (id $($caSummary.EmergencyGroupId))" } else { 'does not exist yet - will be created automatically if this run applies changes' }
            Write-BaselineHost "  Emergency-access group '$($caSummary.EmergencyGroupDisplayName)': $groupState. Excluded from every policy's user condition." 'Yellow'
            Write-BaselineHost "  IMPORTANT: every Conditional Access policy this toolkit creates or updates is set to state=$($caSummary.ReportOnlyState) (REPORT-ONLY). None is ever enabled/enforced by this toolkit - see README.md." 'Yellow'
        }
    }

    switch ($Mode) {

        'Audit' {
            Write-BaselineHost "`nReading current state for $($catalog.Count) control(s)..." 'Cyan'
            $audit = Invoke-BaselineControlAudit -Catalog $catalog

            $snapPath = Get-BaselineTimestampedPath -Directory $BackupPath -Prefix 'backup' -Extension 'json'
            Save-BaselineSnapshot -AuditResults $audit -Path $snapPath -SourceMode Audit -ConfigSchemaVersion $config.schemaVersion | Out-Null
            Write-BaselineHost "Snapshot written: $snapPath" 'Green'

            $reportPathMd = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'audit-report' -Extension 'md'
            Export-BaselineMarkdownReport -AuditResults $audit -Path $reportPathMd -Title 'M365 Baseline Toolkit - Audit Report (app-only)' | Out-Null
            Write-BaselineHost "Markdown report written: $reportPathMd" 'Green'

            if ($IncludeHtmlReport) {
                $reportPathHtml = [System.IO.Path]::ChangeExtension($reportPathMd, 'html')
                Export-BaselineHtmlReport -AuditResults $audit -Path $reportPathHtml -Title 'M365 Baseline Toolkit - Audit Report (app-only)' | Out-Null
                Write-BaselineHost "HTML report written: $reportPathHtml" 'Green'
            }

            $errorCount = @($audit | Where-Object { $_.Error }).Count
            $nonCompliantCount = @($audit | Where-Object { $_.Compliant -eq $false }).Count
            Write-BaselineHost "`nAudit complete. Non-compliant: $nonCompliantCount. Read errors: $errorCount." 'Cyan'
            if ($errorCount -gt 0) {
                Write-BaselineHost "Read errors may indicate a control's cmdlet doesn't support app-only auth for your tenant - see the report's Error column, and re-check via the interactive Invoke-M365Baseline.ps1 if uncertain." 'Yellow'
                $exitCode = 2
            }
        }

        'Apply' {
            Write-BaselineHost "`n[Pre-change] Reading current state (this is your backup) for $($catalog.Count) control(s)..." 'Cyan'
            $preAudit = Invoke-BaselineControlAudit -Catalog $catalog

            $snapPath = Get-BaselineTimestampedPath -Directory $BackupPath -Prefix 'backup' -Extension 'json'
            Save-BaselineSnapshot -AuditResults $preAudit -Path $snapPath -SourceMode Apply-PreChange -ConfigSchemaVersion $config.schemaVersion | Out-Null
            Write-BaselineHost "Pre-change snapshot (backup) written: $snapPath" 'Green'

            $preReportPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'pre-change-report' -Extension 'md'
            Export-BaselineMarkdownReport -AuditResults $preAudit -Path $preReportPath -Title 'M365 Baseline Toolkit - Pre-Change Report (app-only)' | Out-Null
            Write-BaselineHost "Pre-change report written: $preReportPath" 'Green'
            if ($IncludeHtmlReport) {
                Export-BaselineHtmlReport -AuditResults $preAudit -Path ([System.IO.Path]::ChangeExtension($preReportPath, 'html')) -Title 'M365 Baseline Toolkit - Pre-Change Report (app-only)' | Out-Null
            }

            $changeLogPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'changelog' -Extension 'jsonl'
            Write-BaselineHost "`nApplying desired state (change log: $changeLogPath)..." 'Cyan'

            $rawApplyResults = Invoke-BaselineControlApply -Catalog $catalog -AuditResults $preAudit -ChangeLogPath $changeLogPath `
                -WhatIfMode:$WhatIfPreference -ShouldProcessTarget $PSCmdlet -AcknowledgeRisk:$AcknowledgeFederationBlockAll -StopOnError:$StopOnError
            $applyResults = ConvertTo-AppOnlyResultSet -Results $rawApplyResults

            Write-BaselineHost "`n[Post-change] Re-reading state for attempted controls..." 'Cyan'
            $postAudit = Invoke-BaselineControlAudit -Catalog $catalog

            $postSnapPath = Get-BaselineTimestampedPath -Directory $BackupPath -Prefix 'backup-postchange' -Extension 'json'
            Save-BaselineSnapshot -AuditResults $postAudit -Path $postSnapPath -SourceMode Apply-PostChange -ConfigSchemaVersion $config.schemaVersion | Out-Null

            $postReportPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'post-change-report' -Extension 'md'
            Export-BaselineMarkdownReport -AuditResults $postAudit -Path $postReportPath -Title 'M365 Baseline Toolkit - Post-Change Report (app-only)' -ApplyResults $applyResults | Out-Null
            Write-BaselineHost "Post-change report written: $postReportPath" 'Green'
            if ($IncludeHtmlReport) {
                Export-BaselineHtmlReport -AuditResults $postAudit -Path ([System.IO.Path]::ChangeExtension($postReportPath, 'html')) -Title 'M365 Baseline Toolkit - Post-Change Report (app-only)' -ApplyResults $applyResults | Out-Null
            }

            $failed = @($applyResults | Where-Object { $_.Status -eq 'Failed-AppOnlyUnsupported' })
            $succeeded = @($applyResults | Where-Object { $_.Status -eq 'Success' })
            $skippedManual = @($applyResults | Where-Object { $_.Status -eq 'Skipped-Manual' })
            $skippedOk = @($applyResults | Where-Object { $_.Status -eq 'Skipped-AlreadyCompliant' })

            Write-BaselineHost "`nApply complete. Succeeded: $($succeeded.Count). Already compliant: $($skippedOk.Count). Manual/skipped: $($skippedManual.Count). Failed (app-only): $($failed.Count)." 'Cyan'

            if ($skippedManual.Count -gt 0) {
                Write-BaselineHost "`nControls with no automated fix available (no suitable API exists) - change these by hand:" 'Yellow'
                foreach ($m in $skippedManual) { Write-BaselineHost "  $($m.Id): $($m.Message)" 'Yellow' }
            }

            if ($failed.Count -gt 0) {
                Write-BaselineHost "`nFailures (Failed-AppOnlyUnsupported):" 'Red'
                foreach ($f in $failed) { Write-BaselineHost "  $($f.Id): $($f.Message)" 'Red' }
                $exitCode = 3
            }
        }

        'Restore' {
            $snapshot = Import-BaselineSnapshot -Path $BackupFile
            Write-BaselineHost "Loaded snapshot from $BackupFile (captured $($snapshot.capturedAtUtc), mode $($snapshot.mode))." 'Green'

            $changeLogPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'restore-changelog' -Extension 'jsonl'
            Write-BaselineHost "`nRestoring $($snapshot.controls.Count) control(s) to their recorded values (change log: $changeLogPath)..." 'Cyan'

            $rawRestoreResults = Invoke-BaselineControlRestore -Catalog $catalog -SnapshotControls $snapshot.controls -ChangeLogPath $changeLogPath `
                -WhatIfMode:$WhatIfPreference -ShouldProcessTarget $PSCmdlet -StopOnError:$StopOnError
            $restoreResults = ConvertTo-AppOnlyResultSet -Results $rawRestoreResults

            $postAudit = Invoke-BaselineControlAudit -Catalog $catalog
            $postReportPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'restore-report' -Extension 'md'
            Export-BaselineMarkdownReport -AuditResults $postAudit -Path $postReportPath -Title 'M365 Baseline Toolkit - Restore Report (app-only)' -ApplyResults $restoreResults | Out-Null
            Write-BaselineHost "Restore report written: $postReportPath" 'Green'
            if ($IncludeHtmlReport) {
                Export-BaselineHtmlReport -AuditResults $postAudit -Path ([System.IO.Path]::ChangeExtension($postReportPath, 'html')) -Title 'M365 Baseline Toolkit - Restore Report (app-only)' -ApplyResults $restoreResults | Out-Null
            }

            $failed = @($restoreResults | Where-Object { $_.Status -eq 'Failed-AppOnlyUnsupported' })
            $succeeded = @($restoreResults | Where-Object { $_.Status -eq 'Success' })
            Write-BaselineHost "`nRestore complete. Succeeded: $($succeeded.Count). Failed (app-only): $($failed.Count)." 'Cyan'
            if ($failed.Count -gt 0) {
                Write-BaselineHost "`nFailures (Failed-AppOnlyUnsupported):" 'Red'
                foreach ($f in $failed) { Write-BaselineHost "  $($f.Id): $($f.Message)" 'Red' }
                $exitCode = 3
            }
        }
    }
}
catch {
    Write-BaselineHost "`nFATAL: $($_.Exception.Message)" 'Red'
    $exitCode = 1
}
finally {
    if ($KeepConnectionsOpen) {
        if ($connectedServices.Count -gt 0) {
            Write-BaselineHost "`n-KeepConnectionsOpen set: leaving $($connectedServices -join ', ') connected. Close this PowerShell window when you're done to clear the sessions." 'Yellow'
        }
    }
    else {
        foreach ($conn in $connectedServices) {
            Disconnect-BaselineWorkload -Connection $conn
        }
    }
}

exit $exitCode
