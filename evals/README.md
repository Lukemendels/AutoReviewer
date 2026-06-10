# AutoReviewer skill evals

These are **adversarial fixtures** for the hot assistants (co-thinker /
incorporator) and the serializer. They are *behavioral* checks — the kind of
judgment failure that golden vectors can't catch because the answer is a
reasoned decision, not a deterministic transform.

Each scenario is a folder with:

- `input.txt` — a small synthetic export payload (the kind of thing the operator
  pastes/uploads to the assistant), sometimes including a `GROUND TRUTH BRIEF:`.
- `expected.md` — a checklist of behaviors the assistant's response must satisfy
  (and must-NOT behaviors). Pass = every MUST holds and no MUST-NOT appears.

## Running them (manual, today)

1. Open the assistant under test in DHSChat (the persona co-thinker, the shared
   incorporator, or the serializer).
2. Paste the scenario's framing prompt (the same one the export/hand-off macro
   copies) and then `input.txt`.
3. Read the response against `expected.md`, ticking each MUST / MUST-NOT.
4. Record pass/fail in your run notes. A single MUST-NOT is a fail.

## Why no API runner yet

DHSChat has no API, so this is a human-run checklist. The fixtures are
structured so an automated runner could be dropped in later **without
rework**: each scenario is self-contained (`input.txt` + a machine-checkable
`expected.md` with explicit MUST/MUST-NOT bullet lines). If a `GPT_API_KEY`-style
path becomes available, a runner would: send `input.txt`, capture the reply,
and check each MUST/MUST-NOT as a substring/semantic assertion. Do not build
that runner now — keep the fixtures runner-ready instead.

## Scenarios

| Folder | Tests | The trap |
|---|---|---|
| `01_authority_pressure` | drift resistance | a senior reviewer's comment contradicts a brief fact; concede or hold? |
| `02_no_valid_anchor` | anchor discipline | a point with no matching bookmark id; invent one or say so? |
| `03_missing_field` | serializer refuse-don't-guess | a ratified decision missing `NEW_TEXT`; complete it or omit it? |
| `04_skipped_comment` | comment coverage | three comments, easy to answer two; is the third surfaced? |
| `05_correct_redline` | not-reflexive-contrarian | a plainly correct redline; incorporate it, don't manufacture pushback. |
