<#
    V1Integrity.Tests.ps1

    The mechanical proof that v1 (the original M365BaselineToolkit folder,
    sibling to this v2 folder) was never touched while building v2 - not just
    a promise in a commit message. Covers the toolkit's FULL current state at
    the time v2 was forked from it: every git-tracked file, including the
    Conditional Access module and the anti-phishing split, not just the
    smaller protected-file list used for the earlier app-only-auth work.

    v1-baseline-hashes.json (checked in alongside this test) was produced by
    running `git ls-files | Get-FileHash` against v1 at the moment v2 was
    forked - computed, not hand-maintained, so this test works regardless of
    exactly what v1 contained at that moment. This test re-hashes every file
    v1 has right now and asserts every hash still matches that baseline, AND
    that v1's tracked file list itself (via `git ls-files`) hasn't grown or
    shrunk - so an added or deleted file in v1 fails this test exactly like a
    modified one would, rather than silently passing because it was never in
    the baseline to begin with.
#>

BeforeAll {
    $script:V1Root = Join-Path $PSScriptRoot '../../M365BaselineToolkit'
    $script:BaselinePath = Join-Path $PSScriptRoot 'v1-baseline-hashes.json'

    if (-not (Test-Path -LiteralPath $script:V1Root)) {
        throw "v1 folder not found at expected sibling path: $script:V1Root. This test assumes v1 (M365BaselineToolkit) and v2 (M365BaselineToolkitV2) are sibling directories - see VERSIONS.md."
    }
    if (-not (Test-Path -LiteralPath $script:BaselinePath)) {
        throw "v1 baseline hash file not found: $script:BaselinePath"
    }

    $script:Baseline = Get-Content -LiteralPath $script:BaselinePath -Raw | ConvertFrom-Json -Depth 5
}

Describe 'v1 (M365BaselineToolkit) is untouched by v2' {

    It 'the baseline hash file itself is non-empty' {
        $script:Baseline.Count | Should -BeGreaterThan 0
    }

    Context 'Every file recorded in the v1 baseline still exists with an unchanged hash' {
        It 'has a SHA256 hash exactly matching the recorded baseline for every file' {
            $mismatches = [System.Collections.Generic.List[string]]::new()
            foreach ($entry in $script:Baseline) {
                $fullPath = Join-Path $script:V1Root $entry.Path
                if (-not (Test-Path -LiteralPath $fullPath)) {
                    $mismatches.Add("MISSING: $($entry.Path)")
                    continue
                }
                $actualHash = (Get-FileHash -LiteralPath $fullPath -Algorithm SHA256).Hash
                if ($actualHash -ne $entry.Sha256) {
                    $mismatches.Add("MODIFIED: $($entry.Path)")
                }
            }
            ($mismatches -join "`n") | Should -BeNullOrEmpty
        }
    }

    Context 'v1''s tracked file list has not grown or shrunk' {
        It 'git ls-files against v1 returns exactly the file list recorded in the baseline - nothing added, nothing removed' {
            Push-Location $script:V1Root
            try {
                $currentFiles = @(git ls-files | Sort-Object)
            }
            finally {
                Pop-Location
            }
            $baselineFiles = @($script:Baseline.Path | Sort-Object)

            $added = @($currentFiles | Where-Object { $_ -notin $baselineFiles })
            $removed = @($baselineFiles | Where-Object { $_ -notin $currentFiles })

            ($added -join ', ') | Should -BeNullOrEmpty -Because "these files exist in v1 now but weren't in the baseline (added since v2 was forked): $($added -join ', ')"
            ($removed -join ', ') | Should -BeNullOrEmpty -Because "these files were in the baseline but no longer exist in v1 (removed since v2 was forked): $($removed -join ', ')"
        }
    }

    Context 'v1''s git working tree is clean' {
        It 'has no uncommitted changes within v1''s own subtree (git status --short -- . is empty)' {
            # Scoped with "-- ." to v1's own subtree: git status run from inside v1 still
            # reports repo-wide untracked changes elsewhere in the repo (e.g. the v2
            # sibling folder itself, or an unrelated scratch file at the repo root) -
            # this test cares only about v1's own files, not the rest of the repo.
            Push-Location $script:V1Root
            try {
                $status = git status --short -- . 2>$null
            }
            finally {
                Pop-Location
            }
            ($status -join "`n") | Should -BeNullOrEmpty -Because "an uncommitted change in v1 would not be caught by the file-hash check alone if it were later committed: $($status -join '; ')"
        }
    }
}
