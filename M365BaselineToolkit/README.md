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
  That import is also done with `-Global`: `-UseWindowsPowerShell` generates
  local proxy functions for the remoted commands rather than exporting them
  the normal way, and without `-Global` those proxies were only visible
  inside the function that ran the import — not in the separate
  `SharePointOnlineControls.psm1` module that actually calls `Get-SPOTenant`
  and friends, which surfaced as `Get-SPOTenant is not recognized...` even
  right after a successful `Connect-SPOService`.
  `-InstallMissingModules` also knows about this split: PowerShell 7's
  `Install-Module -Scope CurrentUser` and Windows PowerShell 5.1's
  `Install-Module -Scope CurrentUser` write to two different folders
  (`Documents\PowerShell\Modules` vs. `Documents\WindowsPowerShell\Modules`
  on Windows) — installing `Microsoft.Online.SharePoint.PowerShell` the
  normal way from PS7 would leave it invisible to the Windows PowerShell 5.1
  session that actually loads it, surfacing as `...was not loaded because no
  valid module file was found in any module directory` even though
  `-InstallMissingModules` reported success moments earlier. The toolkit
  installs (and checks for) that one module via a real Windows PowerShell
  5.1 process specifically to land it in the right folder.
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
  - **Conditional Access Administrator** (or a role with
    `Policy.ReadWrite.ConditionalAccess` rights), plus enough rights to manage
    the toolkit's placeholder emergency-access group
    (`Group.ReadWrite.All`) and read service principals
    (`Application.Read.All`, used to verify the built-in "Microsoft Azure
    Management" app before referencing it) — all `CA-*` controls. See
    [Conditional Access controls (report-only)](#conditional-access-controls-report-only)
    below before enabling these.
- Required Graph scopes requested on connect: `Policy.ReadWrite.Authorization`,
  `Policy.ReadWrite.AuthenticationMethod`, `Directory.Read.All`,
  `RoleManagement.Read.Directory`, `Organization.Read.All` (license checks -
  see below), `Policy.Read.All`, `Policy.ReadWrite.ConditionalAccess`,
  `Group.ReadWrite.All`, `Application.Read.All`. `Policy.Read.All` is required
  separately from `Policy.ReadWrite.ConditionalAccess` — confirmed against a
  live tenant, `Get-MgIdentityConditionalAccessPolicy` (the read cmdlet, called
  even from `Set-` to look up an existing toolkit-owned policy) fails with
  `[AccessDenied]: required scopes are missing in the token` under
  `Policy.ReadWrite.ConditionalAccess` alone, even fully admin-consented —
  Conditional Access doesn't follow the usual "ReadWrite implies Read" pattern
  other Graph resources do. Requested on every Graph connection regardless of
  which controls are enabled this run — consenting to a scope costs nothing
  on a tenant that can't use the feature behind it (e.g.
  `Policy.ReadWrite.ConditionalAccess` consents fine on Entra ID Free; it's
  actually reading/writing a CA policy that the license gate below prevents).
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

**Re-running in the same window and don't want to sign in again 4 times?**
By default every run disconnects from Graph/Exchange Online/Teams/SharePoint
when it finishes, even on failure, so no session is left authenticated
longer than one run. Add `-KeepConnectionsOpen` to skip that teardown while
you're iterating:
```powershell
./Invoke-M365Baseline.ps1 -Mode Audit -SharePointAdminUrl https://contoso-admin.sharepoint.com -KeepConnectionsOpen
```
The `Connect-*` cmdlets these services provide don't check for an existing
session themselves - each one unconditionally starts a fresh interactive
sign-in whenever it's called, live session or not. So the toolkit tracks
which connections are still live (in a session-global variable, so it
survives the module reloads between separate runs of the script) and skips
calling `Connect-*` again for anything already connected - that's what
actually avoids the repeat prompts, not `-KeepConnectionsOpen` alone. A run
that omits `-KeepConnectionsOpen` still disconnects everything at the end,
including a connection reused from an earlier run in the same window, so the
*next* run after that reconnects from scratch as expected. Close the
PowerShell window when you're actually done to clear the sessions for good.
This flag works the same way in Apply/Restore.

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

If any control has no suitable API to be configured automatically (no `Set-`
action exists, or Microsoft hasn't published a stable way to set it), Apply
prints a dedicated summary of exactly those controls and where to fix them
by hand directly in the console once the run finishes, in addition to
recording them in the post-change report:
```
Controls with no automated fix available (no suitable API exists) - change these by hand:
  EntraID-GlobalAdminCount: Entra admin center > Identity > Roles & administrators > Global Administrator > review and adjust role assignments.
  ...
```

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
no-op, but it's worth checking ahead of time. Note that `EntraID-MfaRegistrationCampaign`'s
`desiredValue.includeTargets` is required by the Graph API (the campaign has
no effect with zero targets) — the seed config uses the documented
`"all_users"` special group id to target everyone; replace it with a specific
group id under `includeTargets` if you'd rather pilot with a subset first.

Two other real-tenant findings worth knowing about, both already fixed in
this toolkit's code but worth being aware of if you're extending it further:
`Update-MgPolicyAuthorizationPolicy` (used by four EntraID controls) has no
`-AuthorizationPolicyId`/Id parameter at all — `authorizationPolicy` is a
singleton, so `-BodyParameter` alone is the only reliable call shape across
SDK versions. And Exchange Online's built-in default anti-phish policy is
actually named `"Office365 AntiPhish Default"`, not `"Default"` like the
other default policies (`ExchangeOnline-AntiPhishingSpoofIntelligence`/
`ExchangeOnline-AntiPhishingMailboxIntelligence` now resolve it dynamically
via each policy's `IsDefault` flag instead of hardcoding either name).

## Conditional Access controls (report-only)

The `CA-*` controls (`workload: "ConditionalAccess"` in config) create or
update a small set of Microsoft-recommended Conditional Access (CA) policies.

**Every policy this toolkit ever creates or updates is set to
`state = "enabledForReportingButNotEnforced"` ("report-only"). No code path in
this toolkit ever enables/enforces a CA policy. Turning any of these on is a
manual, deliberate step you take yourself later, in the Entra admin center,
after reviewing what the report-only sign-in logs show it would have done.**

### Licensing (Tier 1 / Tier 2)

Conditional Access requires Entra ID P1 at minimum; two controls
(`CA-RequireMfaSignInRisk`, `CA-RequirePasswordChangeUserRisk`) use
Identity Protection risk signals and require Entra ID P2. Each config entry
under `workload: "ConditionalAccess"` has a `tier` field:

- **Tier 1** — needs Entra ID P1 (or P2, which is a superset).
- **Tier 2** — needs Entra ID P2 specifically.

The toolkit checks this itself (`Test-TenantServicePlan`, shared with
`ExchangeOnline-AntiPhishingMailboxIntelligence`'s Defender for Office 365
gate) — no license, no attempt:

- **Entra ID Free** — the whole CA module is skipped. Every `CA-*` control
  reports `Skipped-LicenseInsufficient`; nothing is read or written. The
  other five workload modules (EntraID directory settings, Exchange, Teams,
  SharePoint/OneDrive) are unaffected and keep running on Entra ID Free as
  they always have.
- **Entra ID P1** — Tier 1 controls run; Tier 2 controls report
  `Skipped-LicenseInsufficient`.
- **Entra ID P2** — everything runs.

### Idempotency, overlap detection, and the emergency-access group

Every policy this toolkit manages is named with the fixed prefix
`"[M365 Baseline] "` (e.g. `"[M365 Baseline] Require MFA for all users"`) and
matched by exact display name — safe to re-run, same as every other control
in this toolkit. Before *creating* a toolkit-owned policy (never before
updating one that already exists), the toolkit scans every other existing CA
policy in the tenant for a heuristic match (matching grant controls plus a
matching condition, e.g. an existing MFA-for-all-users policy under any
name) and skips creation — reported as `Skipped-PotentialOverlap`, naming the
conflicting policy — rather than risk creating a duplicate/conflicting
policy. If you've reviewed the conflict and still want this toolkit's
report-only policy created alongside it, set `forceCreateDespiteOverlap: true`
on that control in `config/baseline.config.json` and re-run.

A placeholder security group, `"M365 Baseline - Emergency Access Accounts (DO
NOT DELETE)"`, is created (empty) automatically the first time Apply needs it
and excluded from every policy's user condition. Populate it yourself with
your organization's actual break-glass accounts — this toolkit only ensures
the group exists and is wired into every policy; it never adds members to it.

### Restore never deletes a CA policy

`-Mode Restore` works by calling each control's `Set-` function with the
snapshot's recorded value (see [Restore](#restore--push-a-snapshots-recorded-values-back)
below) - but every `CA-*` control's `Set-` function ignores that value
entirely and only ever creates or updates a policy toward the report-only
compliant state; there is no delete/remove code path anywhere in the shared
CA engine. So if a snapshot recorded a `CA-*` control as non-compliant
(policy didn't exist yet at snapshot time) and you later run Restore, it will
not delete a `[M365 Baseline]`-prefixed policy that a subsequent Apply
created - it just sees the live policy is already compliant and no-ops. This
is deliberate, consistent with this module's report-only, never-destructive
design (see the top of this section) - removing a CA policy automatically,
even during Restore, carries real lockout risk if something were misjudged.
To actually remove a toolkit-created policy, do it by hand: Entra admin
center > Protection > Conditional Access, or
`Remove-MgIdentityConditionalAccessPolicy -ConditionalAccessPolicyId <id>`.

### Non-goals

This module deliberately does not include named-location/trusted-IP/
country-based policies, device-compliance or hybrid-join-based policies, or
any mechanism (flagged or otherwise) to auto-promote a report-only policy to
enabled. If you want any of that, build it as a clearly-separate, explicitly
opt-in addition — never modify this module to make report-only optional.

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
    broker. The toolkit connects normally first and only retries with
    `-DisableWAM` (`ExchangeOnlineManagement` 3.7+, if your installed module
    version supports that switch) if this specific crash actually occurs —
    `-DisableWAM` is a strictly worse fallback otherwise, so it's never
    forced on unconditionally.

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

## App-only (certificate) authentication

`Invoke-M365Baseline.AppOnly.ps1` is an alternative entry point that connects
with certificate-based, app-only (client-credentials) authentication instead
of an interactive sign-in. It exists for unattended/scheduled runs, and for
any workstation where interactive sign-in through Windows Account Manager
(WAM) is unreliable (see "If something goes wrong mid-run" above).

**This is a parallel path, not a replacement.** `Invoke-M365Baseline.ps1`
(interactive) is completely unmodified by this addition and behaves exactly
as it always has — nothing about how you use it today changes. The two
scripts share one `config/baseline.config.json`, one JSON schema, and all
five control catalog modules (`EntraIdControls.psm1`,
`ExchangeOnlineControls.psm1`, `TeamsControls.psm1`,
`SharePointOnlineControls.psm1`, `ConditionalAccessControls.psm1`) byte-for-byte
unchanged — every `Get-`/`Set-` function calls whatever service cmdlets it
always called, and those cmdlets don't know or care how the connection they're
using was established. So a best-practice change made by editing
`baseline.config.json` (or adding a new control to a catalog module) applies
to both scripts automatically; only the *connection method* differs between
them. They also write to the same `/reports/` and `/backups/` folders with the
same file-naming convention — a backup taken by one script is a normal input
to `-Mode Restore` on the other, and vice versa, since Restore only needs a
live connection plus the same `Set-` functions, which both scripts provide.

**Why this needed no copy of the orchestration engine.** `BaselineCore.psm1`
already separates "establish a connection" (`Connect-BaselineWorkload`, called
only from `Invoke-M365Baseline.ps1`'s own top-level script body) from "run the
Audit/Apply/Restore loop" (`Invoke-BaselineControlAudit`/`-Apply`/`-Restore`,
which take an already-connected session as a given and never call any
`Connect-*` cmdlet themselves). Because that separation already existed,
`Invoke-M365Baseline.AppOnly.ps1` didn't need a duplicate copy of
`BaselineCore.psm1`'s engine (no `BaselineCoreAppOnly.psm1` was needed) — it
imports and calls the exact same exported functions
(`Import-BaselineConfig`, `Get-BaselineControlCatalog`,
`Get-BaselineConnectionOrder`, `Assert-BaselineRequiredModules`,
`Invoke-BaselineControlAudit`/`-Apply`/`-Restore`, `Save-`/`Import-BaselineSnapshot`,
`Export-BaselineMarkdownReport`/`-HtmlReport`, `Disconnect-BaselineWorkload`)
the interactive script does. Only the connection step itself is swapped, via
`modules/AppOnlyConnections.psm1`'s `Connect-M365BaselineServicesAppOnly`. The
top-level parameter parsing → connect → mode dispatch → report writing
"glue" in `Invoke-M365Baseline.AppOnly.ps1` is necessarily its own copy (that
glue was never an importable function in the interactive script to begin
with), but no control logic, compliance logic, or report/backup logic is
duplicated anywhere.

**No interactive fallback.** `Invoke-M365Baseline.AppOnly.ps1` never falls
back to an interactive prompt for anything. If app-only auth fails to connect
to a service, the whole run fails immediately (see "Troubleshooting a
connection failure" below). If a specific control's cmdlet fails *after* a
successful connection — e.g. because that particular cmdlet or parameter
combination genuinely doesn't support app-only auth for your tenant — Apply
and Restore report it as `Failed-AppOnlyUnsupported` in this script's own
console output and reports, with a note to re-run that specific control via
the interactive `Invoke-M365Baseline.ps1` instead. It is never silently
retried interactively.

### One-time tenant setup

1. **Register an Entra app.** Entra admin center → **Identity → Applications
   → App registrations → New registration**. Single-tenant is sufficient.
   Note the **Application (client) ID** and **Directory (tenant) ID** — these
   are `-AppId`/`-TenantId` below.
2. **Generate a certificate and attach it to the app.**
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=M365BaselineToolkit-AppOnly" `
       -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable `
       -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 `
       -NotAfter (Get-Date).AddYears(2)
   Export-Certificate -Cert $cert -FilePath ./m365-baseline-app-only.cer
   ```
   Upload `m365-baseline-app-only.cer` under the app registration's
   **Certificates & secrets → Certificates** tab. Note `$cert.Thumbprint` —
   that's `-CertificateThumbprint` below (or export a `.pfx` with
   `Export-PfxCertificate` if you'd rather use `-CertificatePath`).
3. **Grant Graph API permissions.** App registration → **API permissions →
   Add a permission → Microsoft Graph → Application permissions**. Add the
   same permissions `Connect-BaselineWorkload` requests interactively (see
   `$script:GraphScopes` in `BaselineCore.psm1`, and the "Required Graph
   scopes" note under Prerequisites above) as their **Application**
   equivalents — Graph exposes an Application-type permission of the same
   name for each of them (`Policy.ReadWrite.Authorization`,
   `Policy.ReadWrite.AuthenticationMethod`, `Directory.Read.All`,
   `RoleManagement.Read.Directory`, `Organization.Read.All`, `Policy.Read.All`,
   `Policy.ReadWrite.ConditionalAccess`, `Group.ReadWrite.All`,
   `Application.Read.All`). Then **Grant admin consent for &lt;tenant&gt;**.
4. **Grant Exchange Online's app-only permission.** Same **API permissions**
   blade → **Add a permission → APIs my organization uses → Office 365
   Exchange Online → Application permissions → `Exchange.ManageAsApp`** →
   grant admin consent. The permission picker also lists
   `Exchange.ManageAsAppV2` — that's a different, newer permission for
   Microsoft's Admin REST API v2.0 endpoint, not for the
   `ExchangeOnlineManagement` PowerShell module `Connect-ExchangeOnline` uses.
   This toolkit needs the plain `Exchange.ManageAsApp` (no `V2` suffix); don't
   grant both.
5. **Assign directory roles the app's service principal separately needs.**
   API permission consent alone is not enough for Exchange Online, Teams, or
   SharePoint — each also requires the app's *service principal* to hold a
   directory role, the same way a human admin account would (confirmed
   against a real tenant - a missing role assignment produces a generic
   "Could not authenticate" failure with no indication that a role, not a
   permission, is what's actually missing):
   - Entra admin center → **Identity → Roles & administrators → Exchange
     Administrator → Add assignments** → search for the app by name → add it.
   - Same for **Teams Administrator**, if any `Teams-*` control is enabled.
   - Same for **SharePoint Administrator**, if any `SharePointOnline-*`
     control is enabled. Also grant the SharePoint API permission itself in
     the same **API permissions** blade → **Add a permission → APIs my
     organization uses → Office 365 SharePoint Online → Application
     permissions → `Sites.FullControl.All`** → admin consent. (Microsoft
     Graph has no equivalent "SharePoint" permission entry for this — it has
     to come from the separate Office 365 SharePoint Online API.)
     `Microsoft.Online.SharePoint.PowerShell` version `16.0.26712.12000` or
     newer is required (app-only support GA'd November 2025); check
     `Get-Module Microsoft.Online.SharePoint.PowerShell -ListAvailable` and
     `Update-Module` if older.

   Graph itself needs no separate directory-role assignment — the Application
   API permissions granted in step 3 are sufficient on their own.

### Certificate custody

- **On-box / scheduled task:** keep the certificate in the local machine or
  service account's `Cert:\CurrentUser\My` (or `Cert:\LocalMachine\My`, with
  `-CertificateStoreLocation LocalMachine`) store, with NTFS permissions
  restricted to the account the scheduled task runs as. Don't export the
  `.pfx` to disk alongside the script.
- **Hosted in Azure (Automation Account, Azure Functions, a VM, etc.):** use
  Azure Key Vault. Retrieve the certificate at runtime with the Key Vault
  SDK/`Az.KeyVault` and pass the resulting `X509Certificate2` object directly
  via `-Certificate`, rather than writing it to a file or the local store at
  all.
- Either way, treat the certificate as a credential: rotate it before
  `-NotAfter`, and revoke/replace it (both in the app registration and
  wherever it's stored) immediately if it may have been exposed.

### Verified cmdlet-compatibility notes

These are what's documented as *not* supported under app-only auth for the
services this toolkit uses. Nothing on either exclusion list is called by
this toolkit's control modules, so Exchange Online and Teams app-only
coverage for this toolkit's specific cmdlets is expected to work as-is.

- **Exchange Online:** app-only auth excludes Microsoft 365 Group management
  cmdlets and the entire Security & Compliance/Purview cmdlet surface.
  Neither category is used by any `ExchangeOnline-*` control in this toolkit
  (anti-spam, anti-phishing, mailbox auditing, DKIM, transport/auto-forwarding,
  SMTP AUTH — all plain Exchange Online configuration cmdlets).
- **Teams:** app-only auth excludes `New-Team`, the
  `*-CsOnlineApplicationInstance` family, `*PolicyPackage*` cmdlets,
  `*-CsTeamsShiftsConnection*`, `*-CsBatchTeamsDeployment*`,
  `Get-/Set-CsTeamsSettingsCustomApp`, and `Get-MultiGeoRegion`. None of these
  are used by any `Teams-*` control in this toolkit (federation, meeting
  defaults, app permission policy, guest access — all tenant-config cmdlets
  outside that exclusion list).
- **SharePoint:** Microsoft has published **no** exclusion list — app-only
  support for `Microsoft.Online.SharePoint.PowerShell` only reached general
  availability in November 2025. Unlike the other two, this toolkit **makes
  no claim, and no test in this repository asserts,** that `Set-SPOTenant`
  and `Set-SPOBrowserIdleSignOut` (the cmdlets `SharePointOnlineControls.psm1`
  uses) actually work under app-only auth. **This is something you need to
  verify empirically against the exact PowerShell and module version you'll
  run this under, before trusting `-Mode Apply` against a production tenant.**
  Run `-Mode Audit` first with only SharePoint controls enabled and confirm
  it reads cleanly; then test `-Mode Apply` against a non-production
  tenant/site collection if you have one.

### Parameters specific to this script

| Parameter | Purpose |
|---|---|
| `-AppId` | The app registration's Application (client) ID. |
| `-TenantId` | Tenant id (GUID) or verified domain, e.g. `contoso.onmicrosoft.com`. |
| `-Organization` | Required if any enabled control needs `ExchangeOnline` — the tenant's `*.onmicrosoft.com` domain specifically (not a GUID). |
| `-SpoAdminUrl` | Required if any enabled control needs `SharePointOnline` — same meaning as the interactive script's `-SharePointAdminUrl`. |
| `-CertificateThumbprint [-CertificateStoreLocation]` | Certificate from a local store (`CurrentUser` default, or `LocalMachine`). |
| `-CertificatePath [-CertificatePassword]` | Certificate from a `.pfx` file. |
| `-Certificate` | A pre-built `X509Certificate2` object (e.g. retrieved from Key Vault by the caller). |

Exactly one of the three certificate-input forms is required per run; all
other parameters (`-Mode`, `-ConfigPath`, `-ReportPath`, `-BackupPath`,
`-BackupFile`, `-InstallMissingModules`, `-StopOnError`,
`-AcknowledgeFederationBlockAll`, `-IncludeHtmlReport`,
`-KeepConnectionsOpen`) are identical in name and meaning to
`Invoke-M365Baseline.ps1`.

```powershell
./Invoke-M365Baseline.AppOnly.ps1 -Mode Audit `
    -AppId <app-id> -TenantId contoso.onmicrosoft.com `
    -CertificateThumbprint <thumbprint>

./Invoke-M365Baseline.AppOnly.ps1 -Mode Apply `
    -AppId <app-id> -TenantId contoso.onmicrosoft.com `
    -Organization contoso.onmicrosoft.com -SpoAdminUrl https://contoso-admin.sharepoint.com `
    -CertificatePath ./app-only.pfx -CertificatePassword (Read-Host -AsSecureString) -WhatIf
```

### Troubleshooting a connection failure

A connection failure names the specific service and the certificate input
form used, then points here. It's always one of three distinct problems:

1. **Certificate problem** — expired, revoked, or the certificate you're
   passing doesn't match the public key actually uploaded to the app
   registration. Check `$cert.NotAfter` and the thumbprint against what's
   listed under the app registration's **Certificates & secrets**.
2. **Missing/unconsented API permission** — the specific service's API
   permission (Graph Application permissions, `Exchange.ManageAsApp`, or
   SharePoint's application permission) wasn't added, or was added but never
   admin-consented. Check **API permissions** on the app registration — an
   unconsented permission shows a warning icon there.
3. **Missing directory role assignment** — Exchange Online, Teams, and
   SharePoint each also require the app's service principal to hold a
   directory role (Exchange Administrator / Teams Administrator / SharePoint
   Administrator respectively), separate from and in addition to API
   permission consent. Check **Roles & administrators** in the Entra admin
   center for the relevant role's **Assignments**. This is the single most
   common cause of a generic "Could not authenticate" SharePoint failure with
   an otherwise fully correct cert/permission setup.

### Tests

`tests/AppOnlyConnections.Tests.ps1` covers `Connect-M365BaselineServicesAppOnly`
(all three certificate-input forms, the uniform-certificate-object guarantee,
lazy connect, and specific error messages), a mechanically-enforced file-hash
check that none of the files listed in "This is a parallel path, not a
replacement" above have changed, and an integration test that runs both
entry-point scripts against the same config and identically-mocked service
responses and diffs their Audit-mode compliance verdicts.
