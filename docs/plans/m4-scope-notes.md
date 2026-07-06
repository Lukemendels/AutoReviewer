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
