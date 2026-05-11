---
name: ria-table-inserter
description: Reviews draft prose for TSA Regulatory Impact Analyses (RIAs) and policy impact analyses to identify where tables should replace narrative quantitative content, and produces formatted markdown tables ready to insert. Use whenever the user has draft RIA or policy analysis prose and needs (1) detection of passages where quantitative data is buried in sentences or bullets, (2) selection of the correct TSA table archetype, or (3) formatted tables ready to drop into the draft. Trigger on phrases like "review my draft," "where should tables go," "format these numbers as a table," "is this a table opportunity," or any request involving RIA prose that contains cost figures, population estimates, PRA burden hours, discount rates, or before/after rule comparisons. Even if the user does not explicitly ask for tables, trigger this skill when reviewing RIA narrative drafts because TSA convention requires tabular presentation of quantitative content.
---

# RIA Table Inserter

## What this skill does

TSA RIAs treat tables as the primary vehicle for quantitative data. Bulleted lists and inline prose enumeration of numbers are explicit anti-patterns under TSA convention and OMB Circular A-4 review expectations. This skill reviews narrative draft prose, identifies passages where quantitative content is buried in sentences or bullets, and produces formatted tables ready to drop in.

## Core principle

If a passage contains two or more parallel quantitative data points across categories, years, or stakeholder groups, it belongs in a table. Prose enumeration of numbers reads as imprecise and creates legal vulnerability under administrative review. Qualitative reasoning, statutory interpretation, and stakeholder feedback synthesis stay in continuous paragraphs — those are not table candidates.

## Workflow

When given a draft, execute these steps in order:

1. Scan the full draft for table triggers (see Trigger Signals).
2. For each trigger, identify which of the three table archetypes applies.
3. Extract the data points from the surrounding prose.
4. Produce a formatted markdown table.
5. Draft a one-sentence prose lead-in to introduce the table and a brief one-to-two sentence interpretation to follow it.
6. Return findings as an inline review: location → original prose → archetype → recommended replacement.

Do not rewrite prose that contains no table trigger. Preserve the third-person objectivity of the surrounding voice ("TSA estimates," "the analysis indicates").

## Trigger signals

A passage is a table candidate when it contains any of the following:

- Multiple parallel cost or population figures across categories (e.g., "$97.6 million for Freight Rail, $120 million for PTPR, $45 million for Pipelines")
- Year-over-year projections of costs, populations, or burdens
- Before-and-after rule comparisons (current standard vs. proposed rule)
- PRA burden calculations with response counts and time-per-response figures
- Discount rate pairings (undiscounted alongside 3 percent and 7 percent)
- Population estimates broken down by stakeholder type
- Forecasted adoption or compliance rates over a temporal horizon
- Any existing bulleted list containing dollar figures, hour counts, or response counts

## Table archetypes

### Archetype 1: Comparative Regulatory Framework

Use when the prose compares the current regulatory state to the proposed rule. Most common in deregulatory actions and significant rule modifications.

Columns: Current Standard | Proposed Rule | Operational Impact | Estimated Cost Savings (or Cost)

Example:

| Current Standard | Proposed Rule | Operational Impact | Estimated Cost Savings |
|---|---|---|---|
| Requires annual renewal of security program | Revises to renewal every three years | Aligns part 1548 renewal period with the TSA-approved Certified Cargo Screening Program (CCSP) | Decreased administrative burden for active IAC population |
| Annual audit and recertification requirement | Triennial audit requirement | Permits regulated entities to redirect resources to supply chain operations | $800,000 annualized |

### Archetype 2: PRA Burden with Lettered Formula Columns

Use when the prose describes Paperwork Reduction Act burdens with response counts and time estimates across a three-year approval cycle. Lettered columns make the mathematical relationships explicit and satisfy OMB review.

Columns: Collection Activity | Year 1 (a) | Year 2 (b) | Year 3 (c) | 3-Year Total (d = a+b+c) | Avg Annual (e = d/3) | Time Per Response in hours (f) | Total Time Burden (g = d × f) | Avg Annual Time Burden (h = g/3)

Example:

| Collection Activity | Year 1 (a) | Year 2 (b) | Year 3 (c) | 3-Yr Total (d) | Avg Annual (e) | Time/Response (f) | Total Burden (g) | Avg Annual Burden (h) |
|---|---|---|---|---|---|---|---|---|
| mDL Waiver Application | 15.0 | 10.0 | 5.0 | 30.0 | 10.0 | 20.0 | 600.0 | 200.0 |
| mDL Waiver Resubmission | 13.5 | 9.0 | 4.5 | 27.0 | 9.0 | 5.0 | 135.0 | 45.0 |
| Total | 28.5 | 19.0 | 9.5 | 57.0 | 19.0 | — | 735.0 | 245.0 |

Always include a Total row. Leave the Time/Response cell in the Total row as an em-dash since it does not aggregate.

### Archetype 3: Cost Distribution by Industry Across Temporal Horizon

Use when the prose distributes costs across multiple regulated stakeholder categories over a standard temporal horizon (10 years for operational/structural rules, 5 years for cyclical fee-setting). This is the workhorse table for any rule affecting more than one industry.

Columns: Year | [Industry 1] ($ thousands) | [Industry 2] ($ thousands) | … | Total Regulated Industries Cost

Example:

| Year | Freight Rail ($ thousands) | PTPR ($ thousands) | OTRB ($ thousands) | Pipelines ($ thousands) | Total Cost |
|---|---|---|---|---|---|
| 1 | $97,652 | $119,996 | $188 | $45,000 | $262,836 |
| 2 | $85,000 | $110,000 | $150 | $40,000 | $235,150 |
| 3 | $80,000 | $105,000 | $145 | $35,000 | $220,145 |

For Executive Summary cost tables, also present undiscounted alongside 3 percent and 7 percent discounted totals. This dual-rate presentation is mandatory per OMB Circular A-4 and cannot be omitted.

## Output format

Return findings as a structured inline review. For each trigger:

---

**Location:** [brief locator, e.g., "Section IV.A, paragraph 3" or quote the first ~6 words of the affected sentence]

**Original prose:** [the passage as drafted]

**Archetype:** [1, 2, or 3, with one-line rationale]

**Recommended replacement:**

[Lead-in sentence introducing the table]

[Formatted markdown table]

[One-to-two sentence interpretation drawing the reader's attention to the most material data point]

---

After all individual findings, close with a brief summary: number of tables recommended, archetype distribution, and any data gaps flagged.

## Handling incomplete data

If a passage contains a table trigger but the prose does not contain the underlying figures (e.g., "costs vary by industry" without specific numbers), do not invent values. Flag the gap explicitly:

**Data gap:** [what's needed, e.g., "Year-by-year cost breakdown for OTRB entities through Year 10. Draft references aggregate but does not enumerate."]

Inventing figures to populate tables creates legal vulnerability under arbitrary-and-capricious review and breaks the audit trail back to source data.

## What not to do

- Do not flag qualitative passages — statutory interpretation, market failure reasoning, stakeholder feedback synthesis, methodological discussion — for tabular conversion. These belong in prose.
- Do not invent figures. Flag data gaps instead.
- Do not recommend tables for single isolated data points or simple inline citations of one figure.
- Do not use bold or italic formatting within table cells.
- Do not modify the surrounding prose voice. Preserve third-person objectivity.
- Do not strip the dual discount rate presentation (undiscounted, 3 percent, 7 percent) from any Executive Summary cost table.
- Do not consolidate distinct stakeholder categories into a single "Other" row to simplify a table. Industry-level granularity is what makes the table defensible.
