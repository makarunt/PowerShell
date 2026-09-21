# M365 Baseline Toolkit — versions in this repository

This repository contains **two** copies of the M365 Baseline Toolkit, as
sibling folders:

```
/M365BaselineToolkit     <- v1 (original, unmodified)
/M365BaselineToolkitV2   <- v2 (current, larger control set)
```

## Which one should I use?

**Use `M365BaselineToolkitV2`** unless you have a specific reason to run the
smaller v1 control set. v2 is a strict superset: every control v1 has, plus
seven new ones, plus three existing ones extended with additional fields, plus
a new workload module. See `M365BaselineToolkitV2/README.md`'s "What's new in
v2" section for the full list.

| | v1 (`/M365BaselineToolkit`) | v2 (`/M365BaselineToolkitV2`) |
|---|---|---|
| Controls | 39 | 46 |
| Workload modules | EntraID, ExchangeOnline, Teams, SharePointOnline, ConditionalAccess | same five, **+ M365AdminCenter** |
| Config file | `M365BaselineToolkit/config/baseline.config.json` | `M365BaselineToolkitV2/config/baseline.config.json` (own file, not shared with v1) |
| Status | Frozen. Not modified as part of building v2 — see "How v1's integrity is guaranteed" below. | Actively developed. |

## Why two folders instead of one toolkit with a version flag?

v2 isn't a drop-in upgrade to the same config file the way, for example, the
Conditional Access module or the anti-phishing control split were (those were
additive changes made directly to the one toolkit that existed at the time).
v2 changes the shape of several existing controls' `desiredValue` and adds a
control (`EntraID-AdminConsentWorkflow`) that hard-fails Apply until you
populate a tenant-specific field. Running v1's config against v2's code, or
v2's config against v1's code, would silently produce wrong behavior rather
than a clean error. Two complete, independent folders — each with its own
entry point, modules, config, schema, and tests — make "which toolkit am I
running, against which config" unambiguous by construction, at the cost of
some duplicated code between the two until v1 is eventually retired.

## How v1's integrity is guaranteed

`M365BaselineToolkitV2/tests/V1Integrity.Tests.ps1` is a Pester test that:

1. Re-hashes (SHA256) every file `../M365BaselineToolkit` (v1) currently has,
   and asserts every hash still matches `v1-baseline-hashes.json` — a
   snapshot taken via `git ls-files | Get-FileHash` at the moment v2 was
   forked from v1 (computed, not hand-maintained).
2. Separately asserts v1's tracked file list (`git ls-files`) hasn't grown or
   shrunk since that snapshot, so an added or deleted file in v1 fails the
   test exactly like a modified one would.
3. Asserts v1's git working tree has no uncommitted changes.

Run it any time to mechanically re-verify v1 hasn't drifted:

```powershell
Invoke-Pester -Path M365BaselineToolkitV2/tests/V1Integrity.Tests.ps1
```

## A note on this document's own scope

v1's own `README.md` (`M365BaselineToolkit/README.md`) is **not** modified to
reference this file or v2 — v1 is frozen exactly as it was, including its
documentation, which is itself part of what "v1 is untouched" means here.
This file lives at the repository root, outside both toolkit folders,
specifically so the versioning story can be documented somewhere without
touching v1 at all. `M365BaselineToolkitV2/README.md` links back to this file.
