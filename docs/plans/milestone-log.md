# Milestone log

Terse, factual record of what landed in each milestone, kept current as milestones merge.
Successor to `m4-scope-notes.md`, which stays as-is for the M4 hand-run/ruling history it
already records; this file starts fresh at M4d.

## M4d — landed to `main` (merge commit `3e840f2`)

Findings F-1..F-6 (see `m4-scope-notes.md`'s hand-run section) addressed:

- F-1 — repair loop wired: a bad paste mid-flow now recovers instead of throwing on every
  subsequent Validate click.
- F-2 — structural fence: `injectEdits` no longer crashes with `splitRun: invalid range` on
  a run containing a soft break/tab/tracked deletion; fenced at load time instead.
- F-3 — annotation fence added (temporary; full fix deferred to M6a's sentinelization).
- F-4 — shipped prompt's D1 whole-paragraph-insertion shape fixed; run-text accumulator
  added.
- F-5 — audit record now carries model/persona provenance.
- F-6 — G2's trailing-newline failure message clarified; header-line derivation
  single-sourced against `validate.js`.

Test count: 357 → 384 at merge.

## M6 (`feat/m6-respond`, pending review-merge)

- M6a — export-side sentinelization (`⟦R n: ...⟧` / `⟦C n: ...⟧` encoding for pre-existing
  revisions/comments) plus the reply-nesting fix in `comments.js`, plus a flow-parameterized
  preflight (`checkPreflight(exported, flowType)`) that drops the annotation fence for
  `run-review` while keeping it enforced for `respond-review`.
- M6b — CriticMarkup-extension reply grammar (`[Cn] {>>...<<}`, `[Rn] {>>[AR:accept|reject]
  ...<<}`) and an exactly-once coverage validator (spec §11 amended).
- M6c — threaded-reply and point-comment injection (`injectResponses` in `inject.js`), the
  Respond tab end-to-end in `src/ui/app.js`.

Test count: 384 → 398 at M6c, then 398 → 401 in this review-fixes pass (Fix 1 + Fix 2 below).

### Review-fixes pass (P0 fixed; P1 guarded; P2 this log)

- **P0 — Run Review sentinel gap.** M6a's flow-parameterized preflight dropped the
  annotation fence for `run-review`, but `loadDocxFromBytes` (`src/ui/load.js`) still passed
  `sentinel: flowType === "respond-review"` — so an annotated document now loading fine in
  Run Review still exported with sentinelization *off*, and its pre-existing comment threads
  rendered as malformed nested CriticMarkup, failing G1 on a byte-perfect echo. Fixed by
  passing `sentinel: true` unconditionally; clean documents are unaffected since sentinels
  only wrap existing annotations.
- **P1 — additive-only guard.** M6 is additive-only by design (`injectResponses` only adds
  comment markers adjacent to a revision, never touches the revision itself) but had no test
  pinning that down. Added; passed against the existing implementation with no source change.
