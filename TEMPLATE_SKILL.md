# AutoReviewer Assistant Templates — the two-temperature split

AutoReviewer's review leg runs **hot → (human ratifies) → cold → zero**
(MKS TSA Profile §7.4). The judgment that *reviews* and the mechanism that
*serializes* must live in **separate context windows**, so there are two
assistant templates, not one:

| Temperature | Assistant | Template | Scope | Job |
|---|---|---|---|---|
| **Hot** (divergent) | Co-thinker | `TEMPLATE_SKILL_COTHINKER.md` | One **per persona** | *Synthetic review:* a **two-turn protocol** — Turn 1 clusters comments/revisions into themes (with postures + counter-cases) and stops for theme rulings; Turn 2 emits the numbered decision blocks. |
| **Hot** (divergent) | Incorporator | `TEMPLATE_SKILL_INCORPORATOR.md` | One, **shared** | *Incorporate feedback:* a synthesis brief, then per-item options + recommendation with counter-case, as a decision packet. Style-agnostic, so not a persona. |
| *(human)* | — | — | — | Ratify on paper: keep / fix / cut. |
| **Cold** (convergent) | Serializer | `TEMPLATE_SKILL_SERIALIZER.md` | One, **shared** | Translate the **ratified** decisions into one fenced ```` ```jsonl ```` block (meta line + edits). `serialize_exactly` — never re-decide. |
| **Zero** | VBA applier | — (code) | — | Strip fences, gate the session token + count + comment coverage, then write the JSONL back to Word as tracked changes. |

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
   `TEMPLATE_SKILL_SERIALIZER.md` into its system prompt, and save its URL via
   the dashboard's **Set Serializer URL** button. Every persona shares it.
2. **Incorporator (once):** create one DHSChat assistant, paste
   `TEMPLATE_SKILL_INCORPORATOR.md` into its system prompt, and save its URL via
   the dashboard's **Set Incorporator URL** button. Shared across all documents;
   used by the "Respond to Review" flow.
3. **Co-thinker (per persona):** the training pipeline (Reduce passes) generates
   the persona's style heuristics; they go into `TEMPLATE_SKILL_COTHINKER.md` at
   the `[INSERT REDUCE PASS 3 OUTPUT HERE]` marker. Create a DHSChat assistant
   from it and save its URL in the **Personas** sheet (`AssistantUrl` column —
   this column holds the persona's *co-thinker* URL).

## Why split at all

Co-resident, the serializer inherits the urge to elaborate and rewrites an edit;
the co-thinker inherits rigidity and stops exploring. The physical separation
keeps each at its proper temperature and puts the human's ratification *between*
them, where the reversibility boundary lives.
