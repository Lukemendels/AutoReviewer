# M4 scope notes — Phase 2 hand-run rulings (2026-07-06)

Decisions from the Phase 2 hand-run and the M3b hotfix session that follows it,
recorded here so M4 planning reads them from the repo rather than from memory.

## Rulings

1. **MOTW / Protected View on downloads.** A downloaded `.docx` opening in Word's
   Protected View, requiring one "Enable Editing" click, is accepted behavior — not a
   defect. Exploring an alternate save path to avoid the Mark-of-the-Web prompt is
   OUT of M4 scope.

2. **Composed prompt must teach the D1 whole-paragraph encoding by example.** Both
   failures below were live, human-improvised-prompt failures during the hand-run —
   the model was never shown the shape it needed to produce. M4's `prompt.js`
   `[HARD CONSTRAINTS]` block must include:
   - A worked example of the D1 whole-paragraph insert shape (the token alone on its
     own line, spec §4) — this is also now enforced structurally: a run-anchored
     ins/sub whose new text contains a raw newline is a hard G1 validation failure
     (see the M3b hotfix PR), so the prompt must show the model the *correct*
     mechanism rather than let it discover the rejection by trial and error.
   - An explicit statement of the preserve-every-byte rule, including the header
     comment lines (already G2-enforced, see issue #10) **and the document's final
     trailing newline** — an improvised prompt that doesn't call this out produces a
     response that's one newline short at EOF and fails G2 for a reason that looks
     unrelated to the model's actual edits.

3. **Repair prompts should be mistake-specific.** A generic "re-emit the entire
   document" instruction wastes a retry when the actual problem is narrow and
   nameable. Repair prompts should quote the specific divergence and name the rule
   that was broken (the new newline-in-insertion validator rule's own message is the
   model for this: it names the mistake and shows the correct D1 construction inline,
   rather than just saying "try again").

4. **Ratify UI visual feedback.** Filed as a hand-run finding: Accept/Reject buttons
   worked (the state machine was already correct) but gave no visible on-screen
   feedback about which decision was currently selected. Fixed in the M3b hotfix PR
   (a `data-decision` attribute on the row plus an `is-selected` class on the active
   button, ~10 lines total) rather than filed as a separate issue — it fit the
   hotfix's own timebox. Either way, this was never an M4 blocker.

5. **UI flow text promises a feature that doesn't exist yet.** The current shell UI's
   copy tells the user to "drop a `.docx`" — drag-and-drop isn't implemented; only
   the file-picker path is. M4 must either reword the copy to match the picker-only
   flow, or implement drag-and-drop as part of M4's file-input work. Undecided which;
   flagging so M4 planning makes the call explicitly rather than shipping the
   mismatch forward again.

## Deferred to Fable (M4 milestone review)

Open design questions surfaced during per-PR (Opus) review, to be ruled on in the single
Fable pass after M4c merges. Each is non-blocking for its phase; listed here so Fable gets
a checklist rather than a scavenger hunt across three PR threads. **Status after M4b:**
item 2 is resolved (see below); item 1 (header echo) is the sole remaining open question.

1. **Header echo — §4 single-source vs Issue #10** (M4a, PR #21 — `prompt.js`,
   `buildHardConstraintsSection`). The export's first three header lines are re-emitted
   (indented) inside `[HARD CONSTRAINTS]` in addition to the verbatim `[DOCUMENT]`
   embedding, so the header appears twice. §4's "no section re-renders any part of it"
   reads absolute; but Issue #10 requires naming the exact lines, and the doc's own Issue
   #10 example achieves that *without* reproducing them. Two conformant readings; PR #21
   shipped the reproduce-it one. No drift risk (runtime-derived from the same source).
   Ruling needed: does single-source dominate (→ describe, don't reproduce; drop the
   `headerLines` derivation) or does the literal-echo aid to G2 justify the exception
   (→ keep, confirm empirically in the hand-run)? In-code flag at the site.

2. **Audit `resolvedAnchor` sourcing — RESOLVED in M4b; no Fable action needed.** Spec §12
   requires a resolved anchor per edit. Ruled during M4b Opus review: **option (a)** — the
   anchor is captured at the app layer as `row.edit.anchor` and written to the audit
   record. This needed **no** change to the frozen `validate.js`: its success path already
   returns `{ ...edit, anchor }` (`resolvedEdits.push` in validate.js), so the resolved
   triple rides through `validate() → ratify rows → audit` unmodified. Implemented in
   `audit.js` (`resolvedAnchor: row.edit.anchor`) and covered by `audit.test.js` (asserts
   the audit anchor mirrors the ratify row, plus a `source.sha256 === sourceMap.docHash`
   consistency guard). Option (b) — deferring to M4c's `resolveEdits` split — was therefore
   unnecessary. Left here for provenance so it isn't reopened.

## M4c step 2 (chunk.js) — decisions surfaced and settled

1. **D1 — chunk-mode `promptVersion` is distinct from single-doc.** Confirmed: single-doc
   `PROMPT_TEMPLATE_VERSION` in `prompt.js` stays `"m4a-2026.07-1"`, unchanged; chunk-mode
   prompts get their own version string (e.g. `"m4c-chunk-2026.07-1"`) once a chunk-mode
   prompt template exists, so audit provenance records which template actually ran. **Not
   yet implemented** — see decision 2 below; `prompt.js` has no chunk-mode template yet, so
   there is nowhere for this version string to live in code this session. Whoever adds the
   chunk preamble/per-chunk prompt builder should apply this ruling then, not re-litigate it.

2. **Session wiring scope — chunk.js + validate.js callers only, no UI this session.**
   `chunk.js` (splitting by top-level heading, offset translation, two-phase composition via
   `validateText`/`resolveEdits`) and its equivalence fuzz oracle (`tests/chunk.test.js`) are
   done. `ui/app.js`, `ui/state.js` (the actual multi-paste chunk-mode flow), and `prompt.js`
   (a chunk-mode prompt builder/preamble, and decision 1's version string) are explicitly
   **deferred** — chosen to keep this session's two commits (validate.js split, then
   chunk.js) focused and independently reviewable rather than also redesigning the paste
   flow's state machine in the same pass. `chunk.js` is fully usable by that future work;
   nothing about it needs to change to be wired in.
