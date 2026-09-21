# M365 Baseline Toolkit — versions in this repository

This repository (`makarunt/PowerShell`) now holds only **v1** of the M365
Baseline Toolkit, in `/M365BaselineToolkit`. v2 was built here as a sibling
folder and later split out into its own standalone branch history, once it
was stable, so the two no longer share a checkout, a config file, or a
running instance.

## Where is v2?

**Use v2** unless you have a specific reason to run the smaller v1 control
set — it's a strict superset: every control v1 has, plus seven new ones,
plus three existing ones extended with additional fields, plus a new
workload module. Get it via:

```
git fetch origin v2
git checkout v2
```

v2's own `README.md` (at that branch's root) has the full "What's new in v2"
list and its own version of this history.

| | v1 (`/M365BaselineToolkit`, this branch) | v2 (`v2` branch, repo root) |
|---|---|---|
| Controls | 39 | 46 |
| Workload modules | EntraID, ExchangeOnline, Teams, SharePointOnline, ConditionalAccess | same five, **+ M365AdminCenter** |
| Config file | `M365BaselineToolkit/config/baseline.config.json` | `config/baseline.config.json` (own file, not shared with v1) |
| Status | Frozen — unmodified since v2 was forked from it. | Actively developed. |

## Why a separate branch instead of one toolkit with a version flag?

v2 isn't a drop-in upgrade to the same config file the way, for example, the
Conditional Access module or the anti-phishing control split were (those were
additive changes made directly to the one toolkit that existed at the time).
v2 changes the shape of several existing controls' `desiredValue` and adds a
control (`EntraID-AdminConsentWorkflow`) that hard-fails Apply until you
populate a tenant-specific field. Running v1's config against v2's code, or
v2's config against v1's code, would silently produce wrong behavior rather
than a clean error. Two complete, independent trees — each with its own
entry point, modules, config, schema, and tests — make "which toolkit am I
running, against which config" unambiguous by construction.

They started as sibling folders in one checkout specifically so v2 could be
built and tested against a known-good, provably-untouched copy of v1 without
either affecting the other. Once v2 was stable, it was extracted (with its
own commit history intact) onto the `v2` branch and removed from this one,
since that sibling-folder arrangement had served its purpose and the two no
longer need to sit in the same checkout to develop independently.
