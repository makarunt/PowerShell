# M365 Baseline Toolkit

A PowerShell toolkit that applies, audits, and can roll back a minimum-viable
security/governance baseline across a Microsoft 365 tenant's Entra ID,
Exchange Online, Teams, OneDrive for Business, and SharePoint Online. It is
idempotent, safe to re-run, and always backs up current state before changing
anything.

## Design summary

**Architecture.** The desired state is a JSON file
(`config/baseline.config.json`) that contains **data only** — every control's
id, whether it's enabled, and its typed desired value. No PowerShell syntax
ever lives in that file and it is never `Invoke-Expression`'d. Each control's
actual logic — how to read the tenant's current value and how to change it —
lives in a pair of versioned PowerShell functions, `Get-<Id>State` and
`Set-<Id>State`, grouped into one module per workload
(`EntraIdControls.psm1`, `ExchangeOnlineControls.psm1`,
`TeamsControls.psm1`, `SharePointOnlineControls.psm1`). `BaselineCore.psm1`
is the orchestrator: it loads/validates the config, matches every config
entry to its `Get-`/`Set-` pair by id (a mismatch in either direction is a
hard validation error, before any connection is made), manages connections,
computes compliance, drives Apply/Restore, and writes reports/backups/change
logs. `Invoke-M365Baseline.ps1` is the single CLI entry point with three
modes.

This means changing a value for next year's guidance is normally a
`baseline.config.json` edit — no code change, no redeploy. Only a genuinely
new control needs a new `Get-`/`Set-` function pair.

**Config file layout.**

```
/M365BaselineToolkit
  Invoke-M365Baseline.ps1          # entry point, all 3 modes
  /config
    baseline.config.json           # desired-state data (edit this to change behavior)
    baseline.config.schema.json    # JSON Schema used to validate it
  /modules
    BaselineCore.psm1              # orchestration, connections, compliance, reports, backup/restore
    EntraIdControls.psm1
    ExchangeOnlineControls.psm1
    TeamsControls.psm1
    SharePointOnlineControls.psm1
  /tests                           # Pester 5 suite, fully mocked, no live tenant needed
  /reports                         # created at runtime, timestamped, never overwritten
  /backups                         # created at runtime, timestamped, never overwritten
```

**Why JSON, not `.psd1`, for the config.** JSON is easy to diff in a pull
request, validates cleanly against a schema with `Test-Json`, and is trivial
for non-PowerShell tooling (CI, a wiki, a GRC system) to read or generate.
`.psd1` would give native comments, but at the cost of a format only
PowerShell tooling parses well. Given this config is meant to be reviewed and
possibly edited outside the PowerShell ecosystem, JSON was the better
trade-off here.

## Prerequisites

- **PowerShell 7+ on Windows** is the primary supported host. `MicrosoftTeams`
  and `Microsoft.Online.SharePoint.PowerShell` have historically had rough
  edges (missing cmdlets, auth quirks) on non-Windows PowerShell — test on
  Windows before relying on this in production elsewhere. Nothing in this
  toolkit's own code is Windows-only; the constraint comes from those two
  vendor modules. `Microsoft.Online.SharePoint.PowerShell` specifically
  targets .NET Framework rather than PowerShell 7's .NET runtime, so even
  **on Windows** the toolkit loads it with `Import-Module -UseWindowsPowerShell`
  (a background Windows PowerShell 5.1 compatibility process) — this is
  handled automatically; you don't need Windows PowerShell 5.1 open
  yourself, just present on the machine, which it is by default on Windows.
- One account with enough admin rights to touch every control below. Either
  **Global Administrator**, or this least-privileged combination:
  - **Exchange Administrator** — all `ExchangeOnline-*` controls and
    `EntraID-UnifiedAuditLog` (which uses an Exchange Online cmdlet despite
    being conceptually an EntraID setting).
  - **Teams Administrator** — all `Teams-*` controls.
  - **SharePoint Administrator** — all `SharePointOnline-*` controls.
  - **Privileged Role Administrator** (or a role with
    `Policy.ReadWrite.Authorization` / `Policy.ReadWrite.AuthenticationMethod`
    rights) — all remaining `EntraID-*` controls, including reading Global
    Administrator role membership.
- Required modules: `Microsoft.Graph` (v2+), `ExchangeOnlineManagement`,
  `MicrosoftTeams` (v6+), `Microsoft.Online.SharePoint.PowerShell`. The
  toolkit checks these at the start of every run and can install missing
  ones for you.

## Installing required modules

```powershell
# Either let the toolkit do it for you on first run:
./Invoke-M365Baseline.ps1 -Mode Audit -InstallMissingModules

# ...or install them yourself ahead of time:
Install-Module Microsoft.Graph -Scope CurrentUser
Install-Module ExchangeOnlineManagement -Scope CurrentUser
Install-Module MicrosoftTeams -Scope CurrentUser
Install-Module Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
```

The toolkit only imports `Microsoft.Graph.Authentication` from the
`Microsoft.Graph` module family (not the whole umbrella package) to keep
Graph connections light, but it checks for `Microsoft.Graph` being installed
per the requirement above.

## Running each mode

All three modes are one script, `Invoke-M365Baseline.ps1`, selected with
`-Mode`.

### Audit — read-only, changes nothing

```powershell
./Invoke-M365Baseline.ps1 -Mode Audit `
    -SharePointAdminUrl https://contoso-admin.sharepoint.com
```

Connects to every workload an enabled control needs, reads current state,
and writes:

- `backups/backup_<UTC timestamp>.json` — a full state snapshot (also usable
  later as a Restore input).
- `reports/audit-report_<UTC timestamp>.md` — one row per control: name,
  workload, current value, desired value, compliant yes/no/unknown,
  automatable, and manual-fix notes where relevant.

Exit code is `0` on a clean audit, `2` if one or more controls errored while
being *read* (e.g. insufficient permissions) — this is distinct from a
control simply being non-compliant, which is an expected finding, not an
error.

### Apply — backs up, then converges non-compliant controls

```powershell
# Dry run first - shows exactly what would change, changes nothing:
./Invoke-M365Baseline.ps1 -Mode Apply -WhatIf `
    -SharePointAdminUrl https://contoso-admin.sharepoint.com

# The real thing:
./Invoke-M365Baseline.ps1 -Mode Apply `
    -SharePointAdminUrl https://contoso-admin.sharepoint.com
```

Apply always does everything Audit does *first* — that pre-change read is
your backup, written before a single setting changes. Then, for every
enabled, automatable, non-compliant control it calls that control's `Set-`
function; already-compliant controls are skipped and logged as
`Skipped-AlreadyCompliant`, and audit-only controls are skipped and logged as
`Skipped-Manual` with the exact GUI location to change them by hand. `-WhatIf`
short-circuits every `Set-` call (via `SupportsShouldProcess`) but still
performs the real pre-change backup/report — local files are never part of
the simulated action, only the tenant-mutating calls are.

Writes, per run: a pre-change snapshot/backup and report, an append-only
change log (`reports/changelog_<timestamp>.jsonl`, one JSON object per
control per attempt: timestamp, id, workload, previous value, attempted
value, result, error message), a post-change snapshot, and a post-change
report showing before/after side by side.

By default, one control failing does not abort the run — every control is
attempted, failures are collected and reported at the end, and the process
exits `3` if anything failed. Pass `-StopOnError` for fail-fast behavior
instead.

**Two controls need tenant-specific config before Apply will run them:**
`ExchangeOnline-DkimSigning` (needs `desiredValue.domains` populated with
your accepted domains) and `Teams-RestrictFederation` (needs
`desiredValue.allowedDomains` populated, unless you deliberately want to
block all external federation). Apply refuses to start — before connecting
to anything — with a specific, actionable error if either is enabled with an
empty required field. If you genuinely want Teams federation fully blocked,
populate `allowedDomains: []` and re-run with
`-AcknowledgeFederationBlockAll`.

### Restore — push a snapshot's recorded values back

```powershell
./Invoke-M365Baseline.ps1 -Mode Restore `
    -BackupFile ./backups/backup_2026-09-17T14-30-00Z.json `
    -SharePointAdminUrl https://contoso-admin.sharepoint.com
```

For every control in the snapshot, calls the same `Set-` function Apply
uses, but targets the value the snapshot recorded as "current at capture
time" — never the live config's `desiredValue`. Same logging, `-WhatIf`
support, and continue-past-failures behavior as Apply. Refuses to run, with
a clear error, if the snapshot file's `snapshotSchemaVersion` doesn't match
what this build of the toolkit understands, rather than guessing at a
partially-understood file.

Note: `ExchangeOnline-DkimSigning`'s `Set-` function always refuses to run
with an empty domain list (there's no safe way to represent "DKIM was never
configured" as an action). If a snapshot recorded an empty domain list for
that control, restoring it will report `Failed` with an explanatory message
rather than silently doing nothing — this is expected, not a bug.

## How the config file works, and how to change it safely

Open `config/baseline.config.json`. Each entry in `controls` looks like:

```json
{
  "id": "SharePointOnline-SharingCapability",
  "workload": "SharePointOnline",
  "enabled": true,
  "automatable": true,
  "desiredValue": "ExternalUserSharingOnly",
  "description": "Tenant-wide external sharing ceiling for SharePoint and OneDrive."
}
```

To change what "compliant" means for a control, edit its `desiredValue` and
re-run Apply. To stop the toolkit from touching a control at all, set
`"enabled": false` — it is then skipped entirely and excluded from reports
and the catalog-matching check. Every `id` must match an implemented
`Get-<id>State`/`Set-<id>State` function pair; the toolkit validates this
before connecting to anything and will tell you exactly which id is missing
its implementation (or which implemented control has no config entry) if
they get out of sync.

Two optional fields change how a control is evaluated, not what it does:

- `complianceMode: "Range"` — for controls like `EntraID-GlobalAdminCount`,
  where `desiredValue` is `{ "min": ..., "max": ... }` and the live value is
  a number that must fall inside that range, instead of matching exactly.
- `requiresPopulatedFields: ["domains"]` — names a property under
  `desiredValue` that Apply refuses to run with an empty value (see the
  DKIM/federation note above).

The config is validated against `config/baseline.config.schema.json` with
`Test-Json`, plus a hand-rolled semantic validator
(`Test-BaselineConfigSemantics` in `BaselineCore.psm1`) for checks JSON
Schema alone can't express clearly (duplicate ids, an `automatable: false`
control missing `manualInstructions`, a `Range` control without `min`/`max`,
etc.). Every validation failure names the specific control id and field.

## `Automatable: false` controls

Three controls in this inventory have no safe or currently-documented
automated remediation. They are always read and reported on in every Audit
(so you can see their current value), but Apply/Restore never attempt to
change them — instead they log `Skipped-Manual` with the exact place to fix
it by hand:

| Control | Where to fix it manually |
|---|---|
| `EntraID-GlobalAdminCount` | Entra admin center → Identity → Roles & administrators → Global Administrator (headcount judgment call; not something to automate) |
| `EntraID-RestrictAdminPortalAccess` | Entra admin center → Identity → Users → User settings → "Restrict access to Microsoft Entra admin center" |
| `EntraID-AdminPasswordResetNotification` | Entra admin center → Identity → Users → User settings |

`EntraID-RestrictAdminPortalAccess` and `EntraID-AdminPasswordResetNotification`
also have no confirmed, stable Microsoft Graph property to *read* as of this
writing, so their Audit report shows `Current: (none)` / `Compliant: Unknown`
rather than a guessed value — the toolkit never fabricates a reading it can't
back with a real API call.

`EntraID-AuthMethodsHardening` and `EntraID-MfaRegistrationCampaign` **are**
implemented as automatable, but Microsoft has changed the nested request-body
shape for `Update-MgPolicyAuthenticationMethodPolicyAuthenticationMethodConfiguration`
and `Update-MgPolicyAuthenticationMethodPolicy` before. Validate the body
parameter shape in this toolkit's `EntraIdControls.psm1` against the
`Microsoft.Graph.Identity.SignIns` module version you have installed before
relying on these two in production — a schema drift here would surface as an
`Update-*` cmdlet error (a `Failed` result in Apply's output), not a silent
no-op, but it's worth checking ahead of time.

## If something goes wrong mid-run

- **Apply was interrupted or a control failed partway through.** Re-running
  Apply is safe: every `Set-` function checks current state and no-ops if
  already compliant, so a second run only touches what's still drifted. You
  do not need to manually figure out what got applied — the change log and
  post-change report from the interrupted run show exactly what happened.
- **You need to undo changes Apply made.** Use Restore with the *pre-change*
  backup file from the run that caused the problem — it's the
  `backup_<timestamp>.json` written right before Apply started changing
  anything (Apply also writes a second `backup-postchange_<timestamp>.json`
  after; don't use that one to undo the run, use the earlier one). List
  `./backups` sorted by time if you're not sure which file corresponds to
  which run — each report also names its paired snapshot's timestamp.
- **A connection fails outright** (wrong role, MFA prompt swallowed by a
  non-interactive session, etc.). The toolkit fails fast with which service
  it couldn't connect to and a pointer to the roles section above; nothing
  is read or changed for any control until every required connection for
  that run succeeds.
- **Graph and Exchange Online connections interfere with each other in the
  same session.** `Microsoft.Graph` and `ExchangeOnlineManagement` are
  [documented as mutually incompatible when both connect interactively in
  one PowerShell process](https://github.com/microsoftgraph/msgraph-sdk-powershell/issues/3576),
  because both bundle their own copy of the MSAL auth library and its Windows
  Account Manager (WAM) broker integration. Two symptoms, both already
  worked around by this toolkit:
  - `Connect-MgGraph` fails with `Method not found: ...WithLogging...` (a
    `MissingMethodException`) — happens when Exchange Online connects
    *first* and pins an older MSAL version. Fixed by always connecting to
    Graph before Exchange Online (`Get-BaselineConnectionOrder` in
    `BaselineCore.psm1`).
  - `Connect-ExchangeOnline` fails with a `NullReferenceException` inside
    `Microsoft.Identity.Client...RuntimeBroker` — happens when Exchange
    Online connects *second*, after Graph has already initialized MSAL's WAM
    broker. Fixed by passing `-DisableWAM` to `Connect-ExchangeOnline`
    (`ExchangeOnlineManagement` 3.7+), which the toolkit does automatically
    when your installed module version supports that switch.

  If you still hit either error after pulling the latest version of this
  toolkit, it's almost always duplicate/stale module versions left behind by
  `Update-Module`. Close **all** PowerShell windows, then run:
  ```powershell
  Get-InstalledModule Microsoft.Graph*, ExchangeOnlineManagement | ForEach-Object {
      Get-InstalledModule $_.Name -AllVersions |
          Sort-Object Version -Descending | Select-Object -Skip 1 |
          ForEach-Object { Uninstall-Module -Name $_.Name -RequiredVersion $_.Version -Force }
  }
  ```
  to remove old duplicates, then open a fresh window and retry.

## Testing

The Pester 5 suite in `/tests` runs entirely against mocked cmdlets — no
live tenant, and none of the four M365 modules, need to be installed to run
it:

```powershell
Install-Module Pester -MinimumVersion 5.5.0 -Scope CurrentUser -Force
Invoke-Pester -Path ./tests
```

- `Config.Tests.ps1` — schema/semantic validation: the shipped config
  passes; a bad enum value, a missing required field, a duplicate id, an
  `automatable: false` control without `manualInstructions`, an unsupported
  `schemaVersion`, and a `Range` control without `min`/`max` all fail with a
  specific error.
- `EntraIdControls.Tests.ps1` / `ExchangeOnlineControls.Tests.ps1` — four
  representative controls' `Get-`/`Set-` functions (`GuestInviteRestriction`,
  `GlobalAdminCount`, `AuthMethodsHardening`, `MailboxAuditingDefault`,
  `DkimSigning`, `DisableSmtpAuth`) exercised with `Mock`, including the
  idempotent no-op path and the DKIM empty-domain-list guard.
- `Orchestrator.Tests.ps1` — compliance diffing (`Compliant`/`NonCompliant`/
  `Unknown` classification, `Range` mode, deep object comparison), that
  `Invoke-BaselineControlApply` correctly classifies
  `Skipped-AlreadyCompliant` / `Skipped-Manual` / `Success` / `Failed`, and
  that Restore mode calls `Set-` with the *snapshot's* recorded value, not
  the live config's `desiredValue`.

**A note on how this was built and verified.** The development sandbox used
to write this toolkit could not reach PowerShellGallery (network policy), so
the Pester suite above could not be executed in that sandbox. Every function
was instead exercised directly with hand-written stub cmdlets standing in
for the Graph/EXO/Teams/SPO surface — including a full run of
`Invoke-M365Baseline.ps1` itself in all three modes against fake versions of
all four required modules — which caught and fixed two real bugs before this
was committed: an empty-array-becomes-`$null`-under-`StrictMode` PowerShell
gotcha that broke the Apply readiness check on the common "nothing to fix"
path, and `-WhatIf` leaking from `$WhatIfPreference` into the report/backup
file writes themselves (which would have meant *no backup file was actually
written* on a `-WhatIf` run, even though the console said it was). Run the
Pester suite in a normal environment with PSGallery access before trusting
this against a production tenant.

## Assumptions and known limitations

- `EntraID-UnifiedAuditLog`'s `workload` is `EntraID` in config/reports (it's
  conceptually a tenant-wide EntraID setting) but its implementation uses
  `Set-AdminAuditLogConfig`, an Exchange Online cmdlet — the toolkit connects
  to Exchange Online for this one control even though it's grouped with
  EntraID everywhere else. This is documented in `BaselineCore.psm1`'s
  `$script:ControlConnectionOverrides`.
- Compliance comparison for object/array-shaped `desiredValue`s is a deep,
  property-order-insensitive structural comparison; array elements are
  compared in order (a domain-list control with the same domains in a
  different order is reported non-compliant — deliberate, since order can
  matter for some of these settings).
- `Teams-RestrictFederation`'s exact `Set-CsTenantFederationConfiguration` /
  `New-CsEdgeAllowList` shape was written against current public docs but
  federation allow-list cmdlets have had several revisions historically;
  verify against your installed `MicrosoftTeams` module version before
  relying on it, same as the two EntraID controls flagged above.
- Out of scope by design (per the original spec): no Conditional Access
  policy management, no Security Defaults toggling, no controls beyond the
  inventory in this README, and no telemetry — this runs entirely against
  your own tenant and stays local.
