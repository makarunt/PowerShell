# M365 Baseline Toolkit (v2)

A PowerShell toolkit that applies, audits, and can roll back a minimum-viable
security/governance baseline across a Microsoft 365 tenant's Entra ID,
Exchange Online, Teams, OneDrive for Business, SharePoint Online, and the
M365 Admin Center's org-wide settings. It is idempotent, safe to re-run, and
always backs up current state before changing anything.

**This is v2.** `M365BaselineToolkitV2` (this folder) and `M365BaselineToolkit`
(v1) are sibling folders at the repository root. v1 is the original,
unmodified toolkit and stays exactly as it was — v2 is a fork of it with a
larger control set, its own `config/baseline.config.json`, and its own tests;
the two do not share a config file or a running instance. See
[`../VERSIONS.md`](../VERSIONS.md) at the repository root for the full
versioning rationale and a side-by-side control-count comparison. If you're
looking for the original, smaller v1 baseline, it's unchanged in
`../M365BaselineToolkit/README.md`.

## What's new in v2

v2 adds seven new controls, extends three existing ones with additional
fields, and introduces a new workload module and two new report statuses.
None of this touches v1 - see `VERSIONS.md`.

**New controls:**

| Control | Workload | Notes |
|---|---|---|
| `EntraID-AdminConsentWorkflow` | EntraID | Requires `desiredValue.reviewers` to be populated before Apply will run - see "`Automatable: false` controls" below for the pattern this follows. |
| `EntraID-GaNotLocalAdminOnJoin` | EntraID | Audit-only: Preview feature, no stable (v1.0) Graph/PowerShell API as of this writing. |
| `M365AdminCenter-SwayExternalSharing` | M365AdminCenter (new module, see below) | Audit-only: no PowerShell/Graph API exists for this setting at all. |
| `SharePointOnline-AzureADB2BIntegration` | SharePointOnline | See "Azure AD B2B integration: a known Microsoft-side deprecation" below - read this before treating a post-apply mismatch as a bug. |
| `SharePointOnline-PreventGuestResharing` | SharePointOnline | Can show a propagation-delay mismatch right after a successful apply - see "Report statuses" below. |
| `SharePointOnline-GuestAccessExpiration` | SharePointOnline | Named guest ACCOUNT expiration - distinct from the pre-existing `SharePointOnline-AnonymousLinkExpiration`, which only governs anonymous "Anyone" links. |
| `SharePointOnline-GuestReauthentication` | SharePointOnline | Email one-time-passcode reauthentication for guests. |

**Extended controls** (same control id, new fields added to `desiredValue`):

| Control | New field(s) |
|---|---|
| `EntraID-AuthMethodsHardening` | `systemCredentialPreferences: { state }` — Microsoft has been gradually rolling out this setting's sign-in-time *effect* tenant-by-tenant through roughly September 2026; a tenant showing this configured but not yet visibly affecting sign-in behavior is not a configuration error. |
| `Teams-BlockConsumerContact` | `externalAccessWithTrialTenants` — Microsoft made `"Blocked"` the tenant-wide default starting July 29, 2024, so this may already read compliant on many tenants; it's still asserted explicitly rather than relying on the inherited default. |
| `Teams-MeetingJoinDefaults` | `allowAnonymousUsersToStartMeeting`, `allowPSTNUsersToBypassLobby` |

**A deliberate deviation from the original v2 gap analysis:** that analysis
asked for `allowedToCreateSecurityGroups` to be folded into
`EntraID-BlockSelfServiceAppCreation`'s existing `Update-MgPolicyAuthorizationPolicy`
call. This toolkit already had a separate, working
`EntraID-BlockSelfServiceSecurityGroupCreation` control for that exact field
before v2 existed. Literally merging it in as asked would have left two
controls independently PATCHing sibling properties of the same
`authorizationPolicy.defaultUserRolePermissions` sub-object within one Apply
run - the same write-interference failure pattern (only the first and last of
several sequential writes to that resource actually persisting) that was
root-caused, though never fully resolved, during this project's now-removed
app-only-authentication work. v2 keeps the two controls separate instead. See
the comment on `Get-EntraID-BlockSelfServiceAppCreationState` in
`EntraIdControls.psm1` for the full reasoning.

### The M365 Admin Center workload

`M365AdminCenterControls.psm1` is a new, dedicated (if currently small)
module for org-wide settings that live under the Microsoft 365 admin center's
"Org settings" rather than any specific workload's own admin center. It's
broken out as its own module - rather than folded into `EntraIdControls.psm1`
or elsewhere - because this workload is likely to gain more controls over
time, and a dedicated module keeps that door open cleanly. It currently has
no live connection of its own: its one control today
(`M365AdminCenter-SwayExternalSharing`) is audit-only with no API at all. Per
the module's own header comment, a future control here that *does* need a
live check most likely goes through Microsoft Graph, the same way
`EntraIdControls.psm1` does - no separate connection path was built
speculatively ahead of an actual need.

### Report statuses: `Skipped-Manual`, `Applied-PendingConfirmation`, `MechanismPossiblyDeprecated`

Every Apply/Restore result carries a `Status`. Three of them matter for
reading a v2 report correctly - `Skipped-Manual` already existed in v1;
`Applied-PendingConfirmation` and `MechanismPossiblyDeprecated` are new in
v2, produced only by a generic post-apply read-back-and-classify helper
(`Test-BaselineApplyOutcome` in `BaselineCore.psm1`), never returned directly
by a control's own `Set-` function:

- **`Skipped-Manual`** — no automated remediation exists for this control
  (`automatable: false` in config). Apply/Restore never call its `Set-`
  function at all; the report shows the exact GUI path from
  `manualInstructions` instead. Unchanged from v1; now also used by the two
  new audit-only controls (`EntraID-GaNotLocalAdminOnJoin`,
  `M365AdminCenter-SwayExternalSharing`).
- **`Applied-PendingConfirmation`** — the `Set-` call itself succeeded (no
  exception), but a follow-up read-back of the control's live value (with one
  short retry) still didn't match the desired value. This is *not* the same
  as `Failed`: `Failed` means the API call itself errored. This status means
  "we sent the change and can't yet confirm it landed" - most often a genuine
  propagation delay on Microsoft's side. Re-audit later; if it's still
  mismatched after a reasonable interval, investigate further.
- **`MechanismPossiblyDeprecated`** — the exact same situation as
  `Applied-PendingConfirmation` above, but for a control explicitly flagged
  `mechanismPossiblyDeprecated: true` in config (currently only
  `SharePointOnline-AzureADB2BIntegration` - see below). The toolkit cannot
  distinguish "still propagating" from "Microsoft has silently disabled this
  mechanism for this tenant" for a setting already known to be affected by
  that kind of migration, so the report says so explicitly rather than
  presenting it as an ordinary compliance failure that looks like a toolkit
  bug.

Both new statuses are produced by the same reusable, control-agnostic
mechanism (`Test-BaselineApplyOutcome`), invoked automatically by
`Invoke-BaselineControlApply`/`Invoke-BaselineControlRestore` after any
control's `Set-` call reports success - not hand-wired per control. They
should never appear for a control whose write-then-read is normally
immediate (which is most of them); seeing one on a control other than the
handful called out in this README is worth a closer look.

### Azure AD B2B integration: a known Microsoft-side deprecation

`SharePointOnline-AzureADB2BIntegration` (`Set-SPOTenant -EnableAzureADB2BIntegration`)
is a real, currently-working control - but Microsoft has been auto-migrating
tenants onto Entra ID (Azure AD) B2B integration since May 2026, on an
unannounced per-tenant schedule. Once a tenant is migrated, the underlying
`Set-SPOTenant` call can become a silent no-op that still reports success,
and the corresponding `Get-SPOTenant` read may or may not reflect the
configured value afterward, depending on how the migration landed for that
tenant. **If a post-apply audit shows this control still non-compliant, that
is not automatically a toolkit bug or a misconfiguration** - it's flagged
`mechanismPossiblyDeprecated: true` in config specifically so Apply reports
this as `MechanismPossiblyDeprecated` (see above) rather than an ordinary
failure. If you see this consistently on a given tenant, that tenant has
likely already been migrated by Microsoft and this setting is effectively
moot there - verify directly in the SharePoint admin center rather than
treating the toolkit's report as the last word.

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
/M365BaselineToolkitV2
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
    ConditionalAccessControls.psm1
    M365AdminCenterControls.psm1   # new in v2 - see "What's new in v2" above
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

Five controls in this inventory (three from v1, two new in v2) have no safe
or currently-documented automated remediation. They are always read and
reported on in every Audit (so you can see their current value), but
Apply/Restore never attempt to change them — instead they log
`Skipped-Manual` with the exact place to fix it by hand:

| Control | Where to fix it manually |
|---|---|
| `EntraID-GlobalAdminCount` | Entra admin center → Identity → Roles & administrators → Global Administrator (headcount judgment call; not something to automate) |
| `EntraID-RestrictAdminPortalAccess` | Entra admin center → Identity → Users → User settings → "Restrict access to Microsoft Entra admin center" |
| `EntraID-AdminPasswordResetNotification` | Entra admin center → Protection → Authentication methods → Password reset → Notifications tab → "Notify all admins when other admins reset their password?" |
| `EntraID-GaNotLocalAdminOnJoin` *(v2)* | Entra admin center → Identity → Devices → Device settings → "Additional local administrators on Microsoft Entra joined devices" — Preview feature, no stable (v1.0) Graph/PowerShell API as of this writing |
| `M365AdminCenter-SwayExternalSharing` *(v2)* | Microsoft 365 admin center → Settings → Org settings → Services → Sway → uncheck "Let people in your organization share their sways with people outside your organization" — no PowerShell/Graph API exists for this setting at all |

`EntraID-RestrictAdminPortalAccess`, `EntraID-GaNotLocalAdminOnJoin`, and
`M365AdminCenter-SwayExternalSharing` have no confirmed, stable API to *read*
as of this writing, so their Audit report shows `Current: (none)` /
`Compliant: Unknown` rather than a guessed value — the toolkit never
fabricates a reading it can't back with a real API call.
`EntraID-AdminPasswordResetNotification` is the same way as of this writing.

**These controls are highlighted everywhere v2 tells you about them - new in
v2** (v1 only ever mixed a `Manual: ...` note into a table cell you'd have to
scroll to find):

- **Console.** Audit, Apply, *and* Restore all print a `⚠ Manual review
  required` block right after the run summary, listing every such control by
  id with its exact fix-it-by-hand instructions, in yellow. Previously only
  Apply printed anything like this, and even then without a warning marker.
- **Markdown report.** A `## ⚠ Manual review required` section appears near
  the top, before the full table, with the same list. Each one's row further
  down in the table also has its Id bolded and prefixed with `⚠`, so it's
  visible while scanning the table alone, not just the summary at the top.
- **HTML report.** The same summary appears as a highlighted callout box, and
  every such control's table row gets an amber background (CSS class
  `manual-row`) plus the same `⚠` prefix on its Id cell.

None of this changes what gets applied or skipped — it's purely about making
sure "this one needs a human" is impossible to miss, whether you're watching
the console, skimming a report's top, or reading the full table.

**The HTML report's table is also fixed-width and responsive**, not just
highlighted. Columns use percentage widths (`table-layout: fixed` with a
`<colgroup>`) instead of the browser default of auto-growing each column to
fit its longest single-line value — which is what previously forced
left-right scrolling to see the whole table, since `Current`/`Desired` often
hold a compact JSON value. Long values now wrap onto multiple lines within
their column instead. A scrollable wrapper around the table is still there as
a safety net for a genuinely unbreakable value (a long token with no spaces),
so only the table scrolls in that rare case, never the whole page.

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
- `EntraIdControls.Tests.ps1` / `ExchangeOnlineControls.Tests.ps1` — a
  representative set of controls' `Get-`/`Set-` functions (`GuestInviteRestriction`,
  `GlobalAdminCount`, `AuthMethodsHardening`, `MailboxAuditingDefault`,
  `DkimSigning`, `DisableSmtpAuth`) exercised with `Mock`, including the
  idempotent no-op path and the DKIM empty-domain-list guard.
  `EntraIdControls.Tests.ps1` also covers all three v2-new/extended EntraID
  controls (`AdminConsentWorkflow` including its empty-reviewers hard-fail,
  `GaNotLocalAdminOnJoin`, the `systemCredentialPreferences` extension to
  `AuthMethodsHardening`) and the `BlockSelfServiceAppCreation`/
  `BlockSelfServiceSecurityGroupCreation` overlap decision above.
- `SharePointOnlineControls.Tests.ps1` / `TeamsControls.Tests.ps1` — new in
  v2: the four new SharePoint controls, the two extended Teams controls (with
  explicit field-preservation assertions confirming the extension didn't
  silently drop a pre-existing field), and an integration test proving the
  `mechanismPossiblyDeprecated` flag routes a persistent post-apply mismatch
  to `MechanismPossiblyDeprecated` for `SharePointOnline-AzureADB2BIntegration`
  and to the ordinary `Applied-PendingConfirmation` for
  `SharePointOnline-PreventGuestResharing` (not flagged).
- `M365AdminCenterControls.Tests.ps1` — new in v2: the audit-only contract
  for `M365AdminCenter-SwayExternalSharing` and its catalog/connection wiring.
- `Orchestrator.Tests.ps1` — compliance diffing (`Compliant`/`NonCompliant`/
  `Unknown` classification, `Range` mode, deep object comparison), that
  `Invoke-BaselineControlApply` correctly classifies
  `Skipped-AlreadyCompliant` / `Skipped-Manual` / `Success` / `Failed`, and
  that Restore mode calls `Set-` with the *snapshot's* recorded value, not
  the live config's `desiredValue`. New in v2: a dedicated block of tests for
  `Test-BaselineApplyOutcome` (the generic read-back-and-classify helper) —
  match-on-retry, persistent mismatch with and without the
  `mechanismPossiblyDeprecated` flag, a genuine thrown `Set-` error staying
  `Failed` and never reaching the helper at all, and StrictMode-safety
  against a hand-built catalog entry with no `MechanismPossiblyDeprecated`
  property. Also new in v2: a block of tests for the manual-review report
  highlighting described under "`Automatable: false` controls" above -
  summary section present/absent, and row-level `⚠`/`manual-row` marking, in
  both the Markdown and HTML report.
- `V1Integrity.Tests.ps1` — new in v2, the mechanical proof that v1 is
  untouched: re-hashes every file `../M365BaselineToolkit` (v1) currently has
  against `v1-baseline-hashes.json` (captured via `git ls-files | Get-FileHash`
  at the moment v2 was forked), and separately asserts v1's tracked file list
  and git working tree haven't changed. See `VERSIONS.md`.

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
