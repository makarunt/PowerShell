#Requires -Version 7.0
<#
.SYNOPSIS
    Applies, audits, or rolls back the M365 Baseline Toolkit's minimum-viable
    security/governance baseline, connecting to Graph and Exchange Online with
    app-only (certificate) authentication instead of interactive sign-in.
.DESCRIPTION
    Identical to Invoke-M365Baseline.ps1 in every way except how it connects to
    Graph and Exchange Online: this script uses the OAuth2 client-credentials
    flow with a certificate (via modules/AppOnlyAuth.psm1) instead of an
    interactive user sign-in, so it never touches Windows Account Manager (WAM)
    and behaves identically on every machine. See APP-ONLY-AUTH.md for the
    one-time Entra app registration, API permission, and certificate setup this
    requires before this script can be used.

    Teams and SharePoint Online connections are NOT covered yet - if an enabled
    control needs one of those, this script still falls back to the normal
    interactive sign-in for that connection only (same code path as
    Invoke-M365Baseline.ps1). See APP-ONLY-AUTH.md's "Known limitations"
    section for why.

    All three run modes (Audit/Apply/Restore), the config file, the control
    catalog, and every control's Get-/Set- logic are identical to
    Invoke-M365Baseline.ps1 - nothing about *what* this toolkit checks or
    changes differs, only *how it authenticates* to Graph and Exchange Online.
.PARAMETER Mode
    'Audit', 'Apply', or 'Restore'. Mutually exclusive run modes.
.PARAMETER TenantId
    Entra tenant id (GUID) or a verified domain (e.g. contoso.onmicrosoft.com).
    Used for the Graph app-only connection.
.PARAMETER ClientId
    The Entra app registration's Application (client) ID. See APP-ONLY-AUTH.md.
.PARAMETER Organization
    The tenant's primary *.onmicrosoft.com domain, used for the Exchange Online
    app-only connection (Exchange Online requires the verified domain, not a
    GUID, even if -TenantId is a GUID). Defaults to -TenantId, so this only
    needs to be passed explicitly when -TenantId is a GUID.
.PARAMETER CertificateThumbprint
    Thumbprint of a certificate already imported into the current user's or
    local machine's certificate store, whose private key matches the public
    certificate uploaded to the app registration. Mutually exclusive with
    -CertificatePath.
.PARAMETER CertificatePath
    Path to a .pfx file containing the certificate and private key, as an
    alternative to installing it in the certificate store first. Mutually
    exclusive with -CertificateThumbprint.
.PARAMETER CertificatePassword
    SecureString password for -CertificatePath. Omit if the .pfx has none.
.PARAMETER ConfigPath
    Path to the desired-state config JSON file.
.PARAMETER SchemaPath
    Path to the config's JSON schema file.
.PARAMETER ReportPath
    Directory reports are written to (created if missing).
.PARAMETER BackupPath
    Directory backup/snapshot files are written to (created if missing).
.PARAMETER BackupFile
    Required for -Mode Restore: path to a snapshot JSON file produced by a
    prior Audit run or Apply run's pre-change phase.
.PARAMETER SharePointAdminUrl
    SharePoint admin center URL (e.g. https://contoso-admin.sharepoint.com).
    Required only when an enabled control needs a SharePointOnline connection
    - that connection is still interactive (see .DESCRIPTION).
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
    of the run. See Invoke-M365Baseline.ps1's help for the full explanation -
    behavior here is identical.
.EXAMPLE
    ./Invoke-M365Baseline-AppOnly.ps1 -Mode Audit -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificateThumbprint <thumbprint>
.EXAMPLE
    ./Invoke-M365Baseline-AppOnly.ps1 -Mode Apply -TenantId contoso.onmicrosoft.com -ClientId <appId> -CertificatePath ./app-only.pfx -CertificatePassword (Read-Host -AsSecureString) -WhatIf
.NOTES
    Exit codes: same as Invoke-M365Baseline.ps1 - 0 success, 1 startup/
    validation/connection failure, 2 Audit completed with read errors,
    3 Apply/Restore completed with control failures.
#>
[CmdletBinding(SupportsShouldProcess, DefaultParameterSetName = 'Thumbprint')]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Audit', 'Apply', 'Restore')]
    [string]$Mode,

    [Parameter(Mandatory)]
    [string]$TenantId,

    [Parameter(Mandatory)]
    [string]$ClientId,

    [Parameter()]
    [string]$Organization = $TenantId,

    [Parameter(Mandatory, ParameterSetName = 'Thumbprint')]
    [string]$CertificateThumbprint,

    [Parameter(Mandatory, ParameterSetName = 'File')]
    [string]$CertificatePath,

    [Parameter(ParameterSetName = 'File')]
    [securestring]$CertificatePassword,

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
    [string]$SharePointAdminUrl,

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
Import-Module (Join-Path $modulesDir 'AppOnlyAuth.psm1') -Force -Global
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

try {
    if ($Mode -eq 'Restore' -and [string]::IsNullOrWhiteSpace($BackupFile)) {
        throw "-Mode Restore requires -BackupFile <path to a snapshot JSON file produced by Audit or Apply>."
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

    # Same union-of-primary-and-extra-connections logic as Invoke-M365Baseline.ps1 -
    # see BaselineCore.psm1's $script:ControlExtraConnections for why.
    $neededConnections = @($catalog | ForEach-Object { @($_.Connection) + @($_.ExtraConnections) })
    $requiredConnections = Get-BaselineConnectionOrder -Connections $neededConnections

    Write-BaselineHost "Checking required modules for: $($requiredConnections -join ', ')" 'Cyan'
    Assert-BaselineRequiredModules -Connections $requiredConnections -InstallMissingModules:$InstallMissingModules

    if (($requiredConnections -contains 'SharePointOnline') -and [string]::IsNullOrWhiteSpace($SharePointAdminUrl)) {
        throw "An enabled control requires a SharePointOnline connection. Pass -SharePointAdminUrl https://<tenant>-admin.sharepoint.com."
    }

    foreach ($conn in $requiredConnections) {
        if (Test-BaselineWorkloadConnected -Connection $conn) {
            Write-BaselineHost "Reusing existing $conn connection (left open by -KeepConnectionsOpen on a prior run in this window)." 'DarkGray'
        }
        else {
            Write-BaselineHost "Connecting to $conn..." 'Cyan'
            switch ($conn) {
                'Graph' {
                    $graphParams = @{ TenantId = $TenantId; ClientId = $ClientId }
                    if ($PSCmdlet.ParameterSetName -eq 'File') {
                        $graphParams['CertificatePath'] = $CertificatePath
                        if ($CertificatePassword) { $graphParams['CertificatePassword'] = $CertificatePassword }
                    }
                    else {
                        $graphParams['CertificateThumbprint'] = $CertificateThumbprint
                    }
                    Connect-BaselineGraphAppOnly @graphParams
                    Set-BaselineWorkloadConnectedState -Connection 'Graph' -Connected $true
                }
                'ExchangeOnline' {
                    $eopAppOnlyParams = @{ Organization = $Organization; ClientId = $ClientId }
                    if ($PSCmdlet.ParameterSetName -eq 'File') {
                        $eopAppOnlyParams['CertificatePath'] = $CertificatePath
                        if ($CertificatePassword) { $eopAppOnlyParams['CertificatePassword'] = $CertificatePassword }
                    }
                    else {
                        $eopAppOnlyParams['CertificateThumbprint'] = $CertificateThumbprint
                    }
                    Connect-BaselineExchangeOnlineAppOnly @eopAppOnlyParams
                    Set-BaselineWorkloadConnectedState -Connection 'ExchangeOnline' -Connected $true
                }
                default {
                    # Teams / SharePointOnline: app-only auth isn't wired up for these yet
                    # (see APP-ONLY-AUTH.md's "Known limitations") - fall back to the same
                    # interactive connection Invoke-M365Baseline.ps1 uses, for this
                    # connection only.
                    Write-BaselineHost "  $conn does not support app-only auth in this script yet - falling back to interactive sign-in for $conn only." 'Yellow'
                    Connect-BaselineWorkload -Connection $conn -SharePointAdminUrl $SharePointAdminUrl
                }
            }
            Write-BaselineHost "Connected to $conn." 'Green'
        }
        # Tracked either way: if THIS run doesn't pass -KeepConnectionsOpen, a
        # connection reused from an earlier run should still be disconnected
        # at the end, same as one this run established itself.
        $connectedServices.Add($conn)
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
            Export-BaselineMarkdownReport -AuditResults $audit -Path $reportPathMd -Title 'M365 Baseline Toolkit - Audit Report' | Out-Null
            Write-BaselineHost "Markdown report written: $reportPathMd" 'Green'

            if ($IncludeHtmlReport) {
                $reportPathHtml = [System.IO.Path]::ChangeExtension($reportPathMd, 'html')
                Export-BaselineHtmlReport -AuditResults $audit -Path $reportPathHtml -Title 'M365 Baseline Toolkit - Audit Report' | Out-Null
                Write-BaselineHost "HTML report written: $reportPathHtml" 'Green'
            }

            $errorCount = @($audit | Where-Object { $_.Error }).Count
            $nonCompliantCount = @($audit | Where-Object { $_.Compliant -eq $false }).Count
            Write-BaselineHost "`nAudit complete. Non-compliant: $nonCompliantCount. Read errors: $errorCount." 'Cyan'
            if ($errorCount -gt 0) { $exitCode = 2 }
        }

        'Apply' {
            Write-BaselineHost "`n[Pre-change] Reading current state (this is your backup) for $($catalog.Count) control(s)..." 'Cyan'
            $preAudit = Invoke-BaselineControlAudit -Catalog $catalog

            $snapPath = Get-BaselineTimestampedPath -Directory $BackupPath -Prefix 'backup' -Extension 'json'
            Save-BaselineSnapshot -AuditResults $preAudit -Path $snapPath -SourceMode Apply-PreChange -ConfigSchemaVersion $config.schemaVersion | Out-Null
            Write-BaselineHost "Pre-change snapshot (backup) written: $snapPath" 'Green'

            $preReportPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'pre-change-report' -Extension 'md'
            Export-BaselineMarkdownReport -AuditResults $preAudit -Path $preReportPath -Title 'M365 Baseline Toolkit - Pre-Change Report' | Out-Null
            Write-BaselineHost "Pre-change report written: $preReportPath" 'Green'
            if ($IncludeHtmlReport) {
                Export-BaselineHtmlReport -AuditResults $preAudit -Path ([System.IO.Path]::ChangeExtension($preReportPath, 'html')) -Title 'M365 Baseline Toolkit - Pre-Change Report' | Out-Null
            }

            $changeLogPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'changelog' -Extension 'jsonl'
            Write-BaselineHost "`nApplying desired state (change log: $changeLogPath)..." 'Cyan'

            $applyResults = Invoke-BaselineControlApply -Catalog $catalog -AuditResults $preAudit -ChangeLogPath $changeLogPath `
                -WhatIfMode:$WhatIfPreference -ShouldProcessTarget $PSCmdlet -AcknowledgeRisk:$AcknowledgeFederationBlockAll -StopOnError:$StopOnError

            Write-BaselineHost "`n[Post-change] Re-reading state for attempted controls..." 'Cyan'
            $postAudit = Invoke-BaselineControlAudit -Catalog $catalog

            $postSnapPath = Get-BaselineTimestampedPath -Directory $BackupPath -Prefix 'backup-postchange' -Extension 'json'
            Save-BaselineSnapshot -AuditResults $postAudit -Path $postSnapPath -SourceMode Apply-PostChange -ConfigSchemaVersion $config.schemaVersion | Out-Null

            $postReportPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'post-change-report' -Extension 'md'
            Export-BaselineMarkdownReport -AuditResults $postAudit -Path $postReportPath -Title 'M365 Baseline Toolkit - Post-Change Report' -ApplyResults $applyResults | Out-Null
            Write-BaselineHost "Post-change report written: $postReportPath" 'Green'
            if ($IncludeHtmlReport) {
                Export-BaselineHtmlReport -AuditResults $postAudit -Path ([System.IO.Path]::ChangeExtension($postReportPath, 'html')) -Title 'M365 Baseline Toolkit - Post-Change Report' -ApplyResults $applyResults | Out-Null
            }

            $failed = @($applyResults | Where-Object { $_.Status -eq 'Failed' })
            $succeeded = @($applyResults | Where-Object { $_.Status -eq 'Success' })
            $skippedManual = @($applyResults | Where-Object { $_.Status -eq 'Skipped-Manual' })
            $skippedOk = @($applyResults | Where-Object { $_.Status -eq 'Skipped-AlreadyCompliant' })

            Write-BaselineHost "`nApply complete. Succeeded: $($succeeded.Count). Already compliant: $($skippedOk.Count). Manual/skipped: $($skippedManual.Count). Failed: $($failed.Count)." 'Cyan'

            if ($skippedManual.Count -gt 0) {
                Write-BaselineHost "`nControls with no automated fix available (no suitable API exists) - change these by hand:" 'Yellow'
                foreach ($m in $skippedManual) { Write-BaselineHost "  $($m.Id): $($m.Message)" 'Yellow' }
            }

            if ($failed.Count -gt 0) {
                Write-BaselineHost "`nFailures:" 'Red'
                foreach ($f in $failed) { Write-BaselineHost "  $($f.Id): $($f.Message)" 'Red' }
                $exitCode = 3
            }
        }

        'Restore' {
            $snapshot = Import-BaselineSnapshot -Path $BackupFile
            Write-BaselineHost "Loaded snapshot from $BackupFile (captured $($snapshot.capturedAtUtc), mode $($snapshot.mode))." 'Green'

            $changeLogPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'restore-changelog' -Extension 'jsonl'
            Write-BaselineHost "`nRestoring $($snapshot.controls.Count) control(s) to their recorded values (change log: $changeLogPath)..." 'Cyan'

            $restoreResults = Invoke-BaselineControlRestore -Catalog $catalog -SnapshotControls $snapshot.controls -ChangeLogPath $changeLogPath `
                -WhatIfMode:$WhatIfPreference -ShouldProcessTarget $PSCmdlet -StopOnError:$StopOnError

            $postAudit = Invoke-BaselineControlAudit -Catalog $catalog
            $postReportPath = Get-BaselineTimestampedPath -Directory $ReportPath -Prefix 'restore-report' -Extension 'md'
            Export-BaselineMarkdownReport -AuditResults $postAudit -Path $postReportPath -Title 'M365 Baseline Toolkit - Restore Report' -ApplyResults $restoreResults | Out-Null
            Write-BaselineHost "Restore report written: $postReportPath" 'Green'
            if ($IncludeHtmlReport) {
                Export-BaselineHtmlReport -AuditResults $postAudit -Path ([System.IO.Path]::ChangeExtension($postReportPath, 'html')) -Title 'M365 Baseline Toolkit - Restore Report' -ApplyResults $restoreResults | Out-Null
            }

            $failed = @($restoreResults | Where-Object { $_.Status -eq 'Failed' })
            $succeeded = @($restoreResults | Where-Object { $_.Status -eq 'Success' })
            Write-BaselineHost "`nRestore complete. Succeeded: $($succeeded.Count). Failed: $($failed.Count)." 'Cyan'
            if ($failed.Count -gt 0) {
                Write-BaselineHost "`nFailures:" 'Red'
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
