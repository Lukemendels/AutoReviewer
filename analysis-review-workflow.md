# Analysis Polish Workflow
## GPT 5.1 / DHSChat — Pre-Ship Review Protocol

---

## Overview

**Goal:** Ship a polished analysis that needs minimal rework and demonstrates independent work quality.

**The model's job in every step is diagnostic, not editorial.** It flags. You decide. You or the model drafts only after you've made a judgment call from the middle seat.

**Gate before any redraft:** Print → review comments → decide what actually needs work → only then return to the context window.

---

## Workflow Steps

```
Step 1 → Context Load (example + draft)
Step 2 → Overview / Structure Audit
Step 3 → Section-by-Section Review
Step 4 → Insert comments into doc
Step 5 → Print & Middle-Seat Decision
Step 6 → Targeted Redrafts (fresh context per section)
Step 7 → Final read-through
```

---

## Step 1 — Context Load

Open a new chat. Paste in this order:

1. **Completed example analysis** (the one you're using as the exemplar)
2. **Your current draft**
3. **The framing prompt below**

### Framing Prompt

```
I am going to ask you to review a draft analysis in multiple steps.

Before I do, here is context you need to do this well:

EXEMPLAR: The first document I've pasted is a completed analysis I've previously written. 
Treat it as the gold standard for structure, argument density, voice, and length calibration. 
Do not reference it explicitly in your output — use it as your internal benchmark.

DRAFT: The second document is my current draft. This is what you are reviewing.

RULES FOR ALL REVIEW STEPS:
- Your job is to flag, not to fix. Do not rewrite anything unless I explicitly ask.
- Be specific. Cite the section and the sentence or passage.
- Do not soften findings. If something is broken, say so plainly.
- Every inch of real estate must serve a unique function. Flag anything that doesn't.
- Arguments, claims, and data points must not repeat across sections. 
  Exceptions: Executive Summary and Summary/Closing sections may preview or recap intentionally. 
  Do not flag those.
- Shorter is better. Flag any sentence that does not add new information.

Confirm you have both documents and are ready for Step 2.
```

---

## Step 2 — Overview / Structure Audit

### Prompt

```
Step 2: Overview and Structure Audit.

Do the following:

1. MAP each section to its unique function in one sentence. 
   If two sections map to the same function, flag it.

2. ASSESS the overall argument arc:
   - Does the analysis open with a clear problem or question?
   - Does each section advance the argument toward a conclusion?
   - Is the conclusion earned by what came before it, or does it introduce new claims?

3. FLAG any structural element present in the exemplar that is absent here 
   and appears load-bearing (i.e., its absence weakens the analysis).

4. FLAG any structural element present in the draft that does not appear 
   in the exemplar and does not clearly justify its existence.

Format your output as:
- SECTION MAP
- ARGUMENT ARC ASSESSMENT
- MISSING ELEMENTS (if any)
- EXTRA ELEMENTS (if any)

Do not rewrite anything.
```

---

## Step 3 — Section-by-Section Review

Run one section at a time. Replace `[SECTION NAME]` each time.

### Prompt

```
Step 3: Section-by-Section Review — [SECTION NAME]

Review only this section. Do the following:

1. LOGIC GAPS: Flag any place where the argument doesn't follow — 
   a claim is made without support, or a conclusion is drawn from insufficient evidence.

2. REDUNDANCY: Flag any argument, claim, or data point in this section 
   that already appeared in a prior section. 
   (Exception: if this is a Summary or Closing section, intentional recaps are allowed.)

3. DEAD WEIGHT: Flag any sentence that does not add new information 
   or does not advance the argument.

4. READER LOSS: Flag any place where a reader unfamiliar with the context 
   would lose the thread — undefined term, unexplained assumption, 
   or a logical jump that needs a bridge.

5. VOICE/LENGTH: Compare this section to the exemplar. 
   Flag any passage that feels longer than it needs to be 
   or diverges noticeably in tone.

Format your output as a numbered list of flags, each with:
- Location (quote the sentence or passage)
- Category (Logic Gap / Redundancy / Dead Weight / Reader Loss / Voice)
- Plain-language explanation of the problem

Do not rewrite anything.
```

---

## Step 4 — Insert Comments

After each section review:

- Copy each flag into a Word comment at the relevant location in your draft
- Label comments by category: `[LOGIC]`, `[REDUNDANT]`, `[DEAD WEIGHT]`, `[CLARITY]`, `[VOICE]`
- Do not make edits yet

---

## Step 5 — Print & Middle-Seat Decision

Print the commented draft.

For each comment, decide one of three things:

| Decision | Meaning |
|---|---|
| **Fix** | The flag is correct. This needs work. |
| **Keep** | You disagree with the flag. Your judgment wins. |
| **Clarify** | You're not sure. Note what context the model may have missed. |

Only the items marked **Fix** go back into the context window.

---

## Step 6 — Targeted Redrafts

**Open a fresh context window for each section you are redrafting.**

Do not reuse the long review thread. Narrow inputs produce targeted outputs.

Load in this order:
1. The exemplar
2. Only the section you are redrafting
3. Your list of Fix flags for that section
4. The prompt below

### Redraft Prompt

```
I need to redraft [SECTION NAME] based on specific flags from a prior review.

EXEMPLAR: [paste exemplar or relevant portion]

CURRENT DRAFT OF THIS SECTION: [paste section]

FLAGS TO ADDRESS:
[paste only the Fix-flagged items for this section]

REDRAFT RULES:
- Match the voice, density, and length calibration of the exemplar.
- Do not introduce new arguments or claims not present in the current draft.
- Do not repeat any argument that appears in another section of the analysis.
- Every sentence must earn its place.
- Shorter is better. If a sentence can be cut without losing meaning, cut it.

Produce one clean redraft of the section only. 
Do not explain your changes.
```

Review the redraft against your original. Take what's better. Keep what's yours.

---

## Step 7 — Final Read-Through

Before you print the final version, one last prompt in a clean context:

### Final Check Prompt

```
Final check before I submit this analysis.

[paste full revised draft]

Do one pass only:

1. Is there any argument or data point that appears more than once 
   outside of the Executive Summary or Summary sections?

2. Is there any sentence that does not add new information?

3. Does the conclusion introduce any new claim not supported by the body?

Return a short list of flags only. If none, say "Clear to ship."
```

---

## Reminders

- The model audits. You decide. The model executes your decision.
- Never ask the model to "improve" or "polish" open-ended. Always give it a specific job.
- Fresh context for every redraft. Long threads drift.
- Your judgment on every Keep decision is not a failure of the process — it's the process working.
