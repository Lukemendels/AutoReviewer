# AutoReviewer

A VBA-orchestrated document-review tool for TSA. It extracts a reviewer's
implicit style from past redlines into a portable **persona**, then uses that
persona to suggest **tracked changes** on new drafts — every suggestion anchored
to a source location, every run leaving an audit trail.

AutoReviewer is the **CAT (Comment Adjudication Tool)** instrument described by
the *MKS TSA Profile v0.2*, which binds the *MKS Normative Core v1.1*. The
governing idea, end to end:

> **Deterministic work is VBA** (parsing, anchoring, serializing, writing back).
> **Judgment is DHSChat** (review and recommendation). Every proposed edit traces
> to a source anchor — no fabrication.

The two spec files (`MKS_TSA_Profile_v0_2.md`, `MKS_Normative_Core_v1.1.md`) are
the standard this tool implements; read them for the *why* behind the design.

## The pipeline

The clipboard, moved by a human, is the transport — there is no DHSChat API.
The review leg runs at four temperatures, with the human's ratification at the
reversibility boundary:

```
  Word doc
     │  ExportWordDocForLLM        (VBA: copy → stamp AR_ anchors → extract text/comments/revisions)
     ▼
  HOT co-thinker assistant         (DHSChat: review; surface recommendation + counter-case; decision packet)
     │  ── human ratifies on paper: keep / fix / cut ──
     ▼
  COLD serializer assistant        (DHSChat: serialize_exactly → strict JSONL; never re-decide)
     │  paste JSONL into LLM_Changes!A8
     ▼
  ApplyWordSuggestionsFromJson     (VBA: validate → write as tracked changes → logic_trace)
     │
     ▼
  Human accepts/rejects in Word, then finalizes (the irreversible step — outside this tool)
```

The tool stops at **tracked-change suggestions** (reversible, rejectable). It
never finalizes or transmits — that irreversible step stays with the human in
Word (Profile §7.2, §8). It also operates on a `*_AR` **working copy**, so the
source of record is never mutated.

## Modules

| Module | Role |
|---|---|
| `modAppCore.bas` | Config (key/value) sheet, Personas registry, sheet setup, styling |
| `modSysUtils.bas` | Clipboard, URL launch, and `ArContentFingerprint` (transport attestation) |
| `modWordUtils.bas` | `StampDocWithArBookmarks` — the `AR_` anchor layer (no-fabrication backbone) |
| `modReviewExport.bas` | Working-copy + stamp + extract + payload fingerprint; hot-prompt and serializer hand-off |
| `modReviewImport.bas` | JSONL → tracked changes (six change types), per-edit Log, JSONL fingerprint |
| `modAudit.bas` | The `Trace` sheet — one `logic_trace` row per run (operator, route, fingerprints) |
| `modDashboardUI.bas` | The dashboard: Train Persona / Run Review / Respond to Review |
| `modTrainingPipeline.bas` | Author filter, corpus builder, three Reduce passes, Save SKILL.md |
| `modSelfTest.bas` | Offline self-test harness: replays `tests/vectors/` against the deterministic VBA (no Word needed) |

The deterministic core (fingerprint, JSONL contract, session gate) has a
**Python reference twin** in `ref/` with golden vectors in `tests/vectors/` —
see `TESTING.md` for the doctrine and the operator round-trip.

Assistant prompts: `TEMPLATE_SKILL.md` (index), `TEMPLATE_SKILL_COTHINKER.md`
(hot, per persona), `TEMPLATE_SKILL_SERIALIZER.md` (cold, shared).

## Getting started

0. **Self-test first.** Import the `.bas` modules, copy `tests\vectors\` next
   to the workbook, and run `RunAllSelfTests` (see `TESTING.md`). The SelfTest
   sheet and `selftest_report.txt` must show `OVERALL: GREEN` before the first
   real document run.
1. Open the `.xlsm` and run `modDashboardUI.BuildDashboard` once. It creates the
   `Config`, `LLM_Changes`, `Personas` (and on first run, `Log` / `Trace`)
   sheets and the dashboard.
2. Set up the shared **Serializer** assistant once: create a DHSChat assistant
   from `TEMPLATE_SKILL_SERIALIZER.md`, then dashboard → **Set Serializer URL**.
3. Train a persona (see `USER_GUIDE.md`) to produce a co-thinker assistant.
4. Run a review: Select Persona → Prepare for Review → ratify → Hand off to
   Serializer → Apply.

Full walkthrough: `USER_GUIDE.md`. Roadmap and module history:
`autoreviewer-v2-scope.md`.

## Notes

- Requires Word + Excel with macros enabled. All COM is late-bound; no library
  references are needed, so it runs on locked-down Office.
- AI **comments/replies** are always authored **"AutoReviewer"** (reliable via
  the object model). AI **insertions/deletions** take their author from
  `Application.UserName`, which the tool sets to "AutoReviewer" — but when Word
  is signed into a Microsoft 365 / DHS account, revision author follows the
  **account** and ignores `UserName`. To force insertion author to "AutoReviewer"
  on account-signed-in Word, check, once: **Word → Options → General → "Always
  use these values regardless of sign in to Office"** (and the tool does the
  rest). The apply step reports, in its summary, the author Word actually
  stamped, so you can see which case you're in. Either way edits are tracked and
  rejectable, and the authored comments are an independent provenance signal.
- Edits are **surgical**: only the changed span is tracked (minimal-diff), so
  changing a word is a one-word revision, not a whole-paragraph rewrite. The
  exported payload is normalized to ASCII punctuation (no "tofu" dashes), and
  internal `AR_` anchor ids are stripped from anything written into the document.
- The `AR_` anchors are stripped as the final apply step, so the delivered
  document is clean and a second pass re-stamps from scratch.
- The `Trace` and `Log` sheets are the defensible artifact — on this substrate
  the audit lineage is the product, not overhead (Profile §1.3). Export writes
  its own Trace row (`Mode = "Export"`), so abandoned reviews leave lineage too.
- **Session binding.** Bookmark ids are generic ordinals, so a stale JSONL
  payload could apply cleanly to the *wrong* document. The serializer's first
  output line is therefore a meta line carrying the export fingerprint and the
  edit count; the apply step verifies both **before opening Word** and refuses
  the whole payload on any mismatch (default-deny, no partial apply). After a
  successful apply, `LLM_Changes!A8:A…` is cleared so stale payloads cannot
  linger. Edits apply in two passes — text/comment changes first, then
  `accept_revision`/`reject_revision` — because revision verdicts can delete
  ranges other edits target (the Log sheet records the pass per line).
- No assistant URLs live in source. Set the persona co-thinker URL in the
  Personas sheet, the shared Serializer/Incorporator URLs via the dashboard,
  and optionally a `CustomGptUrl` Config key as the fallback chat URL.
