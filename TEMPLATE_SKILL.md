# AutoReviewer Assistant System-Prompt Templates — the two-temperature split

AutoReviewer's review leg runs **hot → (human ratifies) → cold → zero**. The
judgment that *reviews* and the mechanism that *serializes* must live in
**separate context windows**, so there are two assistant templates, not one:

| Temperature | Assistant | Template | Scope | Job |
|---|---|---|---|---|
| **Hot** (divergent) | Co-thinker | `TEMPLATE_SKILL_COTHINKER.md` | One **per persona** | *Synthetic review:* a **three-turn protocol** — Turn 1 clusters comments/revisions into themes (with postures + counter-cases) and stops for theme rulings; Turn 2 emits the numbered decision blocks and stops for per-block KEEP/FIX/CUT rulings; Turn 3 folds those rulings into the FINAL RATIFIED PACKET. |
| **Hot** (divergent) | Incorporator | `TEMPLATE_SKILL_INCORPORATOR.md` | One, **shared** | *Incorporate feedback:* a **three-turn protocol** — same structure as the Co-thinker (THEMES / BLOCKS / FINAL RATIFIED PACKET). Style-agnostic, so not a persona. |
| *(human)* | — | — | — | Ratify on paper: keep / fix / cut. |
| *(side)* | Researcher | `TEMPLATE_SKILL_RESEARCHER.md` | One, **shared** | A focused side-investigation: takes a *research brief*, pulls from attached sources, computes per house formulas, and returns **DRAFT + frozen FIGURES + FOOTNOTES**. Spawned when answering a comment needs data not in the doc. |
| *(reference)* | Citation | `TEMPLATE_SKILL_CITATION.md` | One, **shared** + canonical | The house citation/calculation standard. Standalone "I just need a citation" assistant, **and** the source embedded in the Researcher. |
| **Cold** (convergent) | Serializer | `TEMPLATE_SKILL_SERIALIZER.md` | One, **shared** | Translate the **ratified** decisions into one fenced ```` ```jsonl ```` block (meta line + edits). `serialize_exactly` — never re-decide. |
| **Zero** | VBA applier | — (code) | — | Strip fences, gate session token + count + comment coverage, then write the JSONL back to Word as tracked changes (incl. footnote citations). |

Behaviors shared by both hot assistants: **comment coverage** (every comment
gets a block or an explicit `NO_ACTION`, ending with a `COVERAGE:` line); a
**ground-truth brief** slot (facts the operator pastes — agreement is earned by
evidence, not by a redline's author); **audience labeling** (external-facing
comments begin `Program office:`); and **output hygiene** (plain-text, straight
quotes/hyphens, no `AR_` ids in content).

The Co-thinker and the Incorporator are two hot postures on the same pipeline:
the Co-thinker *generates* critique (synthetic review), the Incorporator *reads*
someone else's feedback (use case 2). Both feed the same human ratification and
the same shared Serializer.

## Wiring them up

1. **Serializer (once):** create one DHSChat assistant, paste
   `TEMPLATE_SKILL_SERIALIZER.md` into its Instructions field, and save its URL
   via the dashboard's **Set Serializer URL** button. Every persona shares it.
2. **Incorporator (once):** create one DHSChat assistant, paste
   `TEMPLATE_SKILL_INCORPORATOR.md` into its Instructions field, and save its URL
   via the dashboard's **Set Incorporator URL** button. Shared across all
   documents; used by the "Respond to Review" flow.
3. **Co-thinker (per persona):** the training pipeline (Reduce passes) generates
   the persona's style profile; **Save SKILL.md** assembles it into
   `TEMPLATE_SKILL_COTHINKER.md` at the `[INSERT PERSONA STYLE PROFILE]` marker
   and saves the result as `<persona>_cothinker_assistant.md`. Create a DHSChat
   assistant from that assembled file and save its URL in the **Personas** sheet
   (`AssistantUrl` column — this column holds the persona's *co-thinker* URL).
   Set the `CothinkerTemplatePath` key in the Config sheet to the full path of
   `TEMPLATE_SKILL_COTHINKER.md` before using Save SKILL.md for the first time.
4. **Citation (once):** create an assistant from `TEMPLATE_SKILL_CITATION.md` and
   save its URL via **Set Citation URL**. This file is canonical.
5. **Researcher (once):** assemble `TEMPLATE_SKILL_RESEARCHER.md` by pasting the
   body of `TEMPLATE_SKILL_CITATION.md` (sections 1–3) at its
   `[INSERT TEMPLATE_SKILL_CITATION.md BODY HERE]` marker, create the assistant,
   and save its URL via **Set Researcher URL**. Re-paste the citation body
   whenever the canonical file changes.

## Why split at all

Co-resident, the serializer inherits the urge to elaborate and rewrites an edit;
the co-thinker inherits rigidity and stops exploring. The physical separation
keeps each at its proper temperature and puts the human's ratification *between*
them, where the reversibility boundary lives.
