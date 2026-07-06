# Repo State Audit — 2026-07-06

## Unexpected / flagged

- User manually deleted `claude/m3b-injection-1jeupe` and
  `claude/skill-to-dhschat-refactor-ts444s` before this session ran its own
  verification. Those two branches' claims were **not independently
  re-verified** here (they no longer exist to check). Recorded as skipped,
  not confirmed, below.
- `npm ci` reports 5 vulnerabilities (3 moderate, 1 high, 1 critical) in
  dependencies (dev-only, via vitest toolchain). Not addressed — out of
  scope for this housekeeping session.

## main

- HEAD: `ca3e0930b13086a2bbfc7a845afad4185d0ca3f4` — "M3b: Injection (ooxml/inject.js, comments.js, serialize.js)"
- `npm test`: **208/208 passed** (18 test files)
- `node scripts/roundtrip-fixtures.mjs`: clean — 9 fixtures round-tripped (6 with real injected edits, seed=42)
- `python3 scripts/verify_zip_roundtrip.py fixtures /tmp/ar-zip-roundtrip`: clean — all 9 fixtures PASS, byte-identical elsewhere
- CI freshness gate: `npm run build:html` rebuild of `html/autoreviewer-workbench.html` (155.0 KB) — **byte-identical** to committed copy (`git diff` empty)
- `npm run check:compliance -- html/autoreviewer-workbench.html`: all checks PASS

## Open PRs / issues (GitHub, as of audit)

- Open PRs: **0**
- Open issues: **#10, #15, #16** (matches expected)
  - #10 — M4 prompt.js must require the model to echo the export's metadata header (G2 dependency)
  - #15 — export.js: overlapping comment reply ranges render as nested `{==...==}`, failing G1 on any document with a comment reply
  - #16 — Preflight: reject documents with pre-existing tracked changes (or comment replies) at upload, not mid-flow

## Per-branch verification

| Branch | Claim | Evidence | Outcome | Action taken |
|---|---|---|---|---|
| `claude/m3b-injection-1jeupe` | squash-merged, tree identical to main | N/A — branch deleted by user prior to this session's verification pass | **NOT VERIFIED** (moot: already gone) | none (already deleted) |
| `claude/skill-to-dhschat-refactor-ts444s` | 0 commits ahead of main, stale pointer | N/A — branch deleted by user prior to this session's verification pass | **NOT VERIFIED** (moot: already gone) | none (already deleted) |
| `claude/vba-string-concat-perf-0q34r7` | exactly one unmerged commit (`6933d6e`) adding `SKILL_VBA_STRING_PERF.md`, absent from main | `git rev-list --count origin/main..origin/claude/vba-string-concat-perf-0q34r7` → `1`; `git log origin/main..origin/claude/vba-string-concat-perf-0q34r7 --oneline` → `6933d6e` only; `git cat-file -e origin/main:SKILL_VBA_STRING_PERF.md` → does not exist | **CONFIRMED** | Salvaged `SKILL_VBA_STRING_PERF.md` (192 lines) from `6933d6e` onto this branch. Branch to be deleted after this PR merges. |
| `feat/agent-update-sequential-teardown` | one unmerged commit (`c8336f0`), superseded folder reorg; contains `docs/stickshift-tool-file-pattern.md` (salvage) and `docs/AutoReviewer-HTML-Logic-Spec.md` (superseded by `docs/autoreviewer-workbench-spec.md`, do not salvage) | `git rev-list --count origin/main..origin/feat/agent-update-sequential-teardown` → `1`; `git log ... --oneline` → `c8336f0` only; both docs confirmed absent from main via `git cat-file -e`; `docs/autoreviewer-workbench-spec.md` confirmed present on main; content of `docs/AutoReviewer-HTML-Logic-Spec.md` read directly from `c8336f0` — it is an earlier "distilled logic spec" draft explicitly framed as a precursor to a not-yet-written JS spec, consistent with being superseded | **CONFIRMED** | Salvaged `docs/stickshift-tool-file-pattern.md` (287 lines) from `c8336f0` onto this branch. Did not salvage `docs/AutoReviewer-HTML-Logic-Spec.md`. Branch to be deleted after this PR merges. |

## Files salvaged

- `SKILL_VBA_STRING_PERF.md` — from `claude/vba-string-concat-perf-0q34r7` commit `6933d6e`
- `docs/stickshift-tool-file-pattern.md` — from `feat/agent-update-sequential-teardown` commit `c8336f0`

## Branch deletion status

- `claude/m3b-injection-1jeupe` — deleted (by user, prior to this session)
- `claude/skill-to-dhschat-refactor-ts444s` — deleted (by user, prior to this session)
- `claude/vba-string-concat-perf-0q34r7` — pending deletion after this PR merges
- `feat/agent-update-sequential-teardown` — pending deletion after this PR merges
