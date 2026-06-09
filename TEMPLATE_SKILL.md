# AutoReviewer Assistant Templates — the two-temperature split

AutoReviewer's review leg runs **hot → (human ratifies) → cold → zero**
(MKS TSA Profile §7.4). The judgment that *reviews* and the mechanism that
*serializes* must live in **separate context windows**, so there are two
assistant templates, not one:

| Temperature | Assistant | Template | Scope | Job |
|---|---|---|---|---|
| **Hot** (divergent) | Co-thinker | `TEMPLATE_SKILL_COTHINKER.md` | One per persona | Review against the persona's style; surface each recommendation **with its counter-case**, anchored to a bookmark id; output a human-readable **decision packet** (no JSON). |
| *(human)* | — | — | — | Ratify on paper: keep / fix / cut. |
| **Cold** (convergent) | Serializer | `TEMPLATE_SKILL_SERIALIZER.md` | One, shared | Translate the **ratified** decisions into the strict JSONL edit contract. `serialize_exactly` — never re-decide. |
| **Zero** | VBA applier | — (code) | — | Validate and write the JSONL back to Word as tracked changes. |

## Wiring them up

1. **Serializer (once):** create one DHSChat assistant, paste
   `TEMPLATE_SKILL_SERIALIZER.md` into its system prompt, and save its URL via
   the dashboard's **Set Serializer URL** button. Every persona shares it.
2. **Co-thinker (per persona):** the training pipeline (Reduce passes) generates
   the persona's style heuristics; they go into `TEMPLATE_SKILL_COTHINKER.md` at
   the `[INSERT REDUCE PASS 3 OUTPUT HERE]` marker. Create a DHSChat assistant
   from it and save its URL in the **Personas** sheet (`AssistantUrl` column —
   this column holds the persona's *co-thinker* URL).

## Why split at all

Co-resident, the serializer inherits the urge to elaborate and rewrites an edit;
the co-thinker inherits rigidity and stops exploring. The physical separation
keeps each at its proper temperature and puts the human's ratification *between*
them, where the reversibility boundary lives.
