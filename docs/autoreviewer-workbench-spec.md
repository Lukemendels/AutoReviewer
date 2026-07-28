# AutoReviewer Workbench — Implementation Spec

**Audience:** Claude Code
**Repo:** https://github.com/Lukemendels/AutoReviewer (existing VBA implementation — port the *flows*, not the code)
**Reference implementation:** `redline-to-markdown.html` (working docx→CriticMarkup exporter; port into modules, do not rewrite from scratch)
**Normative reference:** `spec-html-tool-compliance.md` (StickShift HTML Tool Compliance Standard). The built artifact MUST pass its §11 checklist. Where the two conflict, the compliance standard wins on packaging/integration; this spec wins on review-pipeline behavior.
**Status:** v1.1 — approved architecture decisions baked in

---

## 1. Vision

Replace the Excel/VBA AutoReviewer cockpit with a **single self-contained HTML file** ("the Workbench") that runs entirely in the browser with zero dependencies, zero network calls, and zero install footprint. The Workbench handles the full document-review lifecycle:

1. **Train Persona** — wizard that produces a reviewer-persona file in OKF markdown, saved by the user into StickShift (StickShift remains the persona registry; the Workbench is stateless).
2. **Run Review** — export a clean `.docx` to CriticMarkup markdown + source map → package a prompt for the LLM (clipboard, human-in-the-loop) → paste the LLM's response back → validate deterministically → human ratifies each edit → inject accepted edits into the *original* `document.xml` as native Word tracked changes + comments → download the reviewed `.docx`.
3. **Respond to Review** — export a `.docx` that already contains redlines/comments (the existing exporter does this) → LLM drafts comment replies and accept/reject recommendations → inject replies as threaded Word comments → download.

### Architecture decisions (settled — do not revisit)

- **The browser writes the `.docx`.** No VBA, no Word COM, no server. JS reads the ZIP (`DecompressionStream`), edits the XML DOM in memory, re-zips (`CompressionStream` + CRC32), and triggers a download.
- **The hot/cold model split is eliminated.** One LLM turn: the co-thinker emits CriticMarkup directly. The former "cold serializer" role is replaced by a **deterministic validator** (§7). Rationale: any job that can be deterministic must be (gates/axioms principle). Faithful serialization is checkable by string equality, so no second model is needed.
- **CriticMarkup is the sole interchange format** for edits. JSONL edit contracts and `AR_` anchor stamping are retired. Anchoring moves to a **source map** built at export time (§5).
- **Clipboard AHA pattern is preserved.** No API calls. The Workbench composes a prompt with a Copy button; the user pastes into DHSChat and pastes the response back. This is both a governance feature and a stack constraint.
- **Human ratification happens on the artifact.** CriticMarkup is human-readable, so the ratification UI renders each proposed edit with accept/reject controls *before* injection. Rejected edits are never written.

### Non-goals (v1)

- Headers, footers, footnotes, endnotes (body only).
- Programmatically accepting/rejecting *pre-existing* tracked changes (Respond flow adds replies + recommendations; the human clicks Accept/Reject in Word).
- Track *formatting* changes (rPrChange/pPrChange) — count and surface them on export, never author them.
- Persistence beyond the session. No localStorage (artifact constraint and privacy posture); offer a downloadable **session file** instead (§5.4).

---

## 2. Constraints (TSA stack)

- Must run from `file://` in Chromium Edge. No CDN, no fetch, no external fonts, no build-time network assumptions in the artifact.
- `navigator.clipboard.writeText` generally works on `file://` in Chromium but keep the `execCommand("copy")` fallback (already in the reference implementation).
- Single deliverable file: `dist/autoreviewer-workbench.html`. Development happens in ES modules; a build step inlines everything (§9).
- The prompt/response leg goes through DHSChat (GPT-5.1) via clipboard. Assume the model can be prompted to return a fenced markdown block; the parser must extract exactly one fenced block from arbitrary surrounding chatter (§6.3).

---

## 3. Module architecture

```
src/
  zip/
    reader.js        # port unzip/readEntry from reference impl
    writer.js        # NEW: crc32, local headers, central directory, EOCD
  ooxml/
    parse.js         # DOMParser wrappers, NS helpers (port from reference)
    export.js        # docx -> {markdown, sourceMap, lockedRanges, comments}  (port + extend)
    inject.js        # NEW: accepted edits -> mutated document.xml DOM
    comments.js      # NEW: create/extend the six comment parts
    serialize.js     # XMLSerializer + part replacement (never pretty-print)
  criticmarkup/
    grammar.js       # tokenizer for the v1 profile (§4)
    strip.js         # markup -> base text (for the equality gate)
    parse.js         # markup -> ordered edit list with md offsets
  sourcemap.js       # schema + offset resolution (md offset -> body path + run offset)
  validate.js        # gates G1–G5 (§7)
  prompt.js          # versioned prompt templates (§6.2)
  persona.js         # OKF schema, import/export, wizard state
  audit.js           # audit record assembly + SHA-256 (crypto.subtle)
  session.js         # session file save/load
  ui/
    app.js           # flow router (three flows, step components)
    ratify.js        # diff rendering + per-edit accept/reject
    diff.js          # minimal LCS diff for drift display (no dependency)
tests/
fixtures/            # .docx corpus (committed binaries + generation script)
build/               # esbuild + inliner -> dist/autoreviewer-workbench.html
```

Keep `zip` and `ooxml` free of DOM-UI imports so they run in Node for tests. `DOMParser`/`XMLSerializer` must be injected or resolved per environment (browser natives; `@xmldom/xmldom` or `happy-dom` in tests — dev dependency only, never bundled). Node ≥ 18 provides `DecompressionStream`/`CompressionStream` globally.

---

## 4. CriticMarkup profile (v1 grammar)

Support exactly these five constructs; reject anything else:

| Token | Meaning | Injection result |
|---|---|---|
| `{++text++}` | insertion | `<w:ins>` wrapping new run(s) |
| `{--text--}` | deletion | `<w:del>` wrapping split run(s), `w:t`→`w:delText` |
| `{~~old~>new~~}` | substitution | adjacent `<w:del>` + `<w:ins>` |
| `{==text==}{>>comment<<}` | anchored comment | `commentRangeStart/End` + `commentReference` + comment part entry |
| `{>>comment<<}` (bare) | point comment | zero-width anchor at that position |

Rules the validator enforces (§7):

- **No nesting.** No CriticMarkup token may contain another.
- **No block-crossing.** A token must open and close within one markdown block (paragraph, heading, list item, or table cell). Multi-paragraph changes are expressed per paragraph.
- **Whole-paragraph insert:** a line consisting solely of one `{++…++}` token (may contain the exporter's own block syntax, e.g. `{++- new bullet++}`). Injects a new `<w:p>` with the paragraph mark itself flagged inserted (`pPr/rPr/ins`).
- **Whole-paragraph delete:** a line whose entire content is one `{--…--}` token. Wraps all runs in `<w:del>` and flags the paragraph mark deleted (`pPr/rPr/del`) so Word merges paragraphs on accept.
- **Substitution `~>` arrow** is the only arrow form supported.
- `{==…==}` must be immediately followed by `{>>…<<}` (highlight without comment is invalid in the *return* direction; the exporter may still emit it for pre-existing comment ranges).
- Literal `{++`-like sequences in document text: the exporter escapes any pre-existing sequence matching a token opener by inserting a zero-width marker recorded in the source map; the validator treats unescaped collisions as G1 failures. (Expected to be vanishingly rare in TSA regulatory prose; correctness over cleverness.)

**Round-trip semantics for validation:** define `strip(md)` as: `{++x++}`→`""`, `{--x--}`→`x`, `{~~a~>b~~}`→`a`, `{>>x<<}`→`""`, `{==x==}`→`x`. Then the fundamental invariant is:

```
strip(llmResponseMarkdown) === exportedMarkdown     // byte equality after \r\n → \n
```

Any violation is fabrication/drift and blocks the pipeline (G2).

---

## 5. Export leg: markdown + source map

### 5.1 Port the reference exporter

`redline-to-markdown.html` already handles: ZIP reading, rels, comment threads (incl. `commentsExtended` resolution/parents), `w:ins`/`w:del`/substitution collapsing, headings, lists, tables, hyperlinks, images, bold/italic, author annotation. Port it into `ooxml/export.js` unchanged in behavior, then extend with the source map.

### 5.2 Source map schema

The exporter must record, for every character of emitted markdown, whether it is **document text** (exists in a `w:t`) or **synthetic** (markdown syntax the exporter invented: `#` heading prefixes, `- ` bullets, `**`/`*` emphasis, `[label](url)` scaffolding, table pipes/dashes, escape backslashes, `[image: …]` placeholders, blank separator lines, CriticMarkup emitted for *pre-existing* redlines).

```js
{
  docHash: "sha256-…",            // of the original .docx bytes
  blocks: [{
    mdStart, mdEnd,               // [start, end) offsets in exported markdown
    kind: "p"|"heading"|"list"|"tableCell",
    bodyPath: [7]                  // index into body.children; tables: [tblIdx, rowIdx, cellIdx, pIdx]
    runs: [{ mdStart, mdEnd, runIndex, charOffset }],  // document-text spans only
  }],
  synthetic: [[start, end), …],   // sorted, non-overlapping
  locked:    [[start, end), …],   // protected islands, see 5.3
}
```

Resolution function (in `sourcemap.js`): given an md offset range, return `{bodyPath, runIndex, charStart, charEnd}` triples, or throw a typed error naming the overlap (synthetic/locked) for the validator to surface.

**Edit-boundary snapping:** if an edit span's boundary falls on synthetic characters (e.g., the LLM deleted bold text including the `**`), snap boundaries inward past synthetic chars. An edit whose *entire* content is synthetic, or that would delete part of a locked island, is rejected (G4).

### 5.3 Locked ranges (protected islands)

Content that is lossy in markdown must be exported as a visible placeholder, recorded as locked, and declared untouchable in the prompt:

- Fields (`w:fldSimple`, `w:instrText` runs) → `⟦field: PAGE⟧` (export the cached display text where available, placeholder otherwise)
- Content controls (`w:sdt`) → export inner text but lock the range in v1
- Images → `[image: name]` (already emitted; now also locked)
- Math (`m:oMath`), embedded objects → `⟦object⟧`, locked

The prompt (§6.2) instructs the model: *text inside ⟦…⟧ and [image: …] must be returned verbatim; do not edit, move, or comment inside it.* The validator enforces this regardless.

### 5.4 Session file

Because the round trip happens over the clipboard, the user may close the tab between copy and paste. Provide **Download session** / **Resume session** (a `.json` containing exported markdown, source map, base64 of the original docx, persona id, prompt version, timestamp). This substitutes for localStorage, which must not be used.

---

## 6. The LLM leg (clipboard AHA)

### 6.1 Flow

1. User drops the `.docx`, picks a persona file (drag the OKF `.md` from StickShift).
2. Workbench renders the composed prompt with a **Copy prompt** button and shows token-estimate + chunking advice (§6.4).
3. User pastes into DHSChat, copies the model's reply, pastes into the Workbench's **Response** pane.
4. Validator runs automatically on paste (§7); on pass, the ratification UI opens (§8).

### 6.2 Prompt template (versioned, embedded, editable in an "advanced" panel)

Structure — assemble in `prompt.js`, record `promptVersion` in the audit:

```
[PERSONA]           ← full OKF persona body
[TASK]              ← flow-specific instructions
[CRITICMARKUP RULES]← the exact v1 grammar from §4, with 2–3 worked examples, e.g.:
                      original: The rule shall apply to all carriers.
                      edited:   The rule {~~shall~>must~~} apply to all {--air--} carriers.
                      {==all carriers==}{>>Confirm scope includes indirect air carriers.<<}
[HARD CONSTRAINTS]  ← return the ENTIRE document; change nothing outside CriticMarkup
                      tokens; never touch ⟦…⟧ or [image: …]; no nesting; no block-crossing;
                      return output inside ONE fenced block ```markdown … ```
[DOCUMENT]          ← the exported markdown, inside its own fence
```

### 6.3 Response envelope

Parser extracts the **last** fenced block from the pasted response (models often preface with chatter). If zero or >1 plausible document-sized fences, show an explicit picker rather than guessing.

### 6.4 Long documents (chunk mode)

Full-document round trip is v1 default (strongest validation: global byte equality). If exported markdown exceeds a configurable threshold (~12k words), offer **chunk by top-level heading**: each chunk gets its own prompt/paste cycle and its own G2 gate; source-map offsets are chunk-relative with a stored base offset. Ratification and injection operate on the merged edit list.

---

## 7. Validation gates (deterministic; each failure blocks with a specific, actionable message)

- **G1 — Grammar.** Tokenize the response. Balanced delimiters, known tokens only, no nesting, no block-crossing, `==` always paired with `>>`.
- **G2 — Fidelity (the fabrication gate).** `strip(response) === exportedMarkdown` byte-for-byte (after newline normalization only — no whitespace forgiveness). On failure, render an LCS word-diff of the two base texts highlighting the drift, with a "copy repair prompt" button that asks the model to re-emit without paraphrasing. This gate is what makes the single-model design safe.
- **G3 — Anchor resolution.** Every edit span resolves through the source map to concrete run offsets after boundary snapping.
- **G4 — Protection.** No edit overlaps a locked range; no edit is entirely synthetic.
- **G5 — Sanity report.** Counts (ins/del/sub/comments), largest single deletion flagged if > N words (configurable, default 50), duplicate-comment detection. G5 warns; G1–G4 block.

Property to encode in tests: for any valid response, `inject → accept-all → extract text` equals `applyEdits(strip⁻¹)` — i.e., accepting everything in Word yields exactly the text the model proposed.

---

## 8. Ratification UI

One row per edit, in document order: context snippet (±80 chars, edit highlighted), type badge, the change rendered as inline diff, and Accept / Reject toggles (default: accept for `sub/ins/del`, always-accept not offered — the human must scroll the full list before **Inject** enables). Bulk controls: accept all, reject all, jump-to-next. Comments are individually toggleable too.

On **Inject**: only accepted edits are written; rejected ones are recorded in the audit with `decision: "rejected"`. This is the human-ratification boundary — nothing reaches the docx without passing through this screen.

---

## 9. Injection + ZIP write-back

### 9.1 XML mutation (`ooxml/inject.js`)

Process edits **sorted by position, right-to-left within each paragraph** so earlier offsets stay valid.

1. Resolve `bodyPath` → `w:p`; compute run-local offsets from the source map.
2. **Split runs** at span boundaries: clone the `w:r`, deep-copy its `rPr`, split the `w:t` text; set `xml:space="preserve"` on any `w:t`/`w:delText` with leading/trailing whitespace.
3. **Deletion:** move the split run(s) into a new `<w:del w:id w:author w:date>`; rename each `w:t` → `w:delText`.
4. **Insertion:** new `w:r` (copy `rPr` from the left-neighbor run so formatting continues naturally) inside `<w:ins …>`.
5. **Substitution:** `w:del` then `w:ins`, adjacent.
6. **Comments:** allocate `w:id = max(existing)+1`; insert `commentRangeStart`/`commentRangeEnd` around the resolved range and a `w:r>w:commentReference` after the end marker. Point comments anchor both markers at the same position.
7. **Whole-paragraph ops** per §4 (paragraph-mark `ins`/`del` in `pPr/rPr`).
8. Revision ids: `w:id` unique document-wide (single counter across ins/del). `w:author` = configurable, default `AutoReviewer — {personaName}` (machine authorship must be visible in the Review pane — attribution is a governance requirement, not a nicety). `w:date` = ISO timestamp.

### 9.2 Comment parts

Word comments span up to six cross-linked parts: `comments.xml`, `commentsExtended.xml` (threading/`done`), `commentsIds.xml`, `commentsExtensible.xml`, the document rels entries, and `[Content_Types].xml` overrides. Implement `ooxml/comments.js` to create-or-extend all of them; missing any linkage is the classic cause of Word's repair prompt. Replies (Respond flow) set `w15:paraIdParent` in `commentsExtended.xml`; generate fresh `w14:paraId` values (8-hex, unique).

### 9.3 Serialization + ZIP writer

- Serialize only mutated parts with `XMLSerializer`. **Never pretty-print or reformat** — whitespace inside `w:t` is content, and reindentation is a repair-prompt generator.
- `zip/writer.js`: for every entry in the original archive, copy untouched parts **byte-for-byte compressed as-is** (reuse original compressed data + original CRC — no recompression needed); for mutated parts, deflate via `new CompressionStream("deflate-raw")`, compute CRC32 (table-driven, ~15 lines), write local file headers, central directory, EOCD. No ZIP64 needed at these sizes; assert total < 4 GB and entry count < 65535.
- Preserve original entry order and names. Output filename: `{original} — reviewed.docx`.

### 9.4 Definition of "it worked"

The downloaded file opens in Word with **no repair prompt**; every accepted edit appears in the Review pane attributed to the configured author; Accept All yields exactly the model's proposed text; comments show correct threading and anchors.

---

## 10. Personas (Train Persona flow)

Personas live in StickShift as OKF markdown; the Workbench only creates and consumes the files. Schema:

```markdown
---
okf: persona
name: RIA Economist Reviewer
description: >-        # SOLE discovery surface — one sentence, trigger-oriented
  Reviews regulatory impact analyses for economic soundness, cost-benefit
  completeness, and OMB Circular A-4 compliance.
version: 1.0
updated: 2026-07-01
---
## Role and voice
## Review priorities        (ordered)
## Style exemplars          (2–5 before/after pairs — these do most of the work)
## Do-not-touch rules       (compiled into [HARD CONSTRAINTS] at prompt time)
## Comment conventions      (when to comment vs. edit; tone)
```

The wizard walks the same conceptual steps as the VBA `Train Persona` column (read the module in the repo and mirror its step semantics; where a step exists only to work around Excel, drop it). Each step is a form; the final step previews the assembled markdown and offers **Download persona.md** with instructions to save into the StickShift skills directory. Also support **exemplar mining**: drop a previously human-reviewed `.docx` and the exporter converts its redlines into candidate before/after exemplar pairs the user can select into the persona.

---

## 11. Respond to Review flow

1. Drop a `.docx` containing existing redlines/comments. Export (the reference impl already renders these as CriticMarkup + threaded `{>>…<<}`), assigning each comment thread and each tracked change a stable label: `[C1]…`, `[R1]…`, included in the markdown.
2. Prompt asks the model to return a **structured response block** (not a full document):

```
[C3] reply: Agreed — revised in §4.2. | recommend: resolve
[R7] recommend: accept | rationale: aligns with A-4 terminology
```

3. Validator: every label must exist; every label addressed at most once; free-text length caps.
4. Ratification UI as in §8.
5. Injection: replies become threaded comments (`paraIdParent` per §9.2) anchored to the same ranges; recommendations on tracked changes become **point comments** adjacent to the change (v1 does not auto-accept/reject — the human does that in Word; v2 may add it).

---

## 12. Audit trail

Every injection produces an audit record: original file SHA-256, output SHA-256, timestamp, persona name+version, prompt version, response SHA-256, full edit list with per-edit decisions (accepted/rejected) and resolved anchors, gate results, app version. v1: offered as a downloadable `.json` sidecar next to the docx (auto-download alongside). v2: additionally embed as a proper custom XML part (requires content-type + rels wiring — do not drop a bare file into the ZIP; unknown unregistered parts can trigger repair).

---

## 13. Testing & CI

- **Runner:** Vitest, Node ≥ 18 (native `CompressionStream`). XML via `@xmldom/xmldom` or `happy-dom` (dev-only).
- **Fixtures:** `fixtures/` holds small committed `.docx` binaries plus `fixtures/generate.py` (python-docx) documenting how each was made: plain paragraphs, headings, lists, tables, hyperlinks, images, bold/italic runs, pre-existing tracked changes, threaded comments (incl. resolved), fields, content controls, a 50-page stressor.
- **Invariant suites:**
  1. *Export determinism:* same input → identical markdown + map.
  2. *No-op round trip:* inject with zero accepted edits → output ZIP's `document.xml` semantically identical; file opens (validate structure with a schema-lite check: balanced tags, required parts present, content-types complete).
  3. *Accept-all equivalence:* for generated random valid edit sets — inject → programmatically apply accept-all (transform `w:ins` unwrap, `w:del` remove, `delText` drop) → extract text → must equal `applyEdits(exportedMarkdown)`.
  4. *Gate properties:* fuzz the response with paraphrase-drift, nesting, block-crossing, locked-range violations; every corruption must be caught by the named gate.
  5. *ZIP writer:* output readable by an independent implementation (Node's `yauzl` or Python `zipfile` in a CI step); CRCs verify.
- **CI:** GitHub Actions — lint, unit, invariants, then build and assert `dist/autoreviewer-workbench.html` is single-file (no `src=`/`href=` to external resources, no `fetch(`), size budget < 400 KB.

---

## 14. Build

esbuild bundles `src/ui/app.js` (which imports everything) to one IIFE; a small script injects the bundle + CSS into `template.html` → `dist/autoreviewer-workbench.html`. No minification tricks that hinder debugging (keep `--sourcemap=inline` in dev builds, strip in release). The dist file is the only deployment artifact — it travels by email/SharePoint like any document.

### 14.1 StickShift compliance scaffolding (per the compliance standard)

`template.html` must carry, and the build must preserve verbatim:

- **Embedded companion skill:** `<script type="text/markdown" id="stickshift-skill" data-skill-slug="autoreviewer-workbench">…</script>` — StickShift skill frontmatter (`type: Skill`, `title`, `description`, `tags`; no `status`), keyword-rich description covering all three flows (tracked changes, redlines, review responses, personas), and a body instructing the assistant to emit the `<HTML_OPEN>` block with `tool: autoreviewer-workbench.html` and `include: - skills/autoreviewer-workbench.md`. Source of truth lives at `src/stickshift/skill.md`; the build inlines it.
- **Identity declaration:** `const STICKSHIFT_TOOL = { file: "autoreviewer-workbench.html", skillSlug: "autoreviewer-workbench", title: "AutoReviewer Workbench" };` — must match the skill block and the `<HTML_OPEN> tool:` line.
- **Onboarding panel:** non-blocking "Set up with StickShift" entry per the standard's §7 (intro, Yes/No branch, Copy-skill button reading from `#stickshift-skill`).
- **Clipboard:** all copy actions (prompt copy, skill copy, repair-prompt copy) route through the standard's §9 `copyText` with the hidden-textarea `execCommand` fallback — the async clipboard API is assumed to FAIL on `file://`; the fallback is the primary path, not insurance.

CI additions: assert `#stickshift-skill` exists with valid frontmatter, `data-skill-slug === STICKSHIFT_TOOL.skillSlug`, `STICKSHIFT_TOOL.file` matches the dist filename and the skill's `tool:` line, and every checklist item in the compliance standard's §11 that is statically checkable.

---

## 15. Milestones & acceptance criteria

**M0 — Scaffold.** Repo layout, build to single HTML incl. StickShift compliance scaffolding (§14.1), CI green, fixture generator. ✅ when `dist` opens on `file://`, shows the shell UI, and passes the compliance standard's §11 checklist.

**M1 — Export module.** Reference impl ported to modules + source map + synthetic/locked tracking + tests. ✅ when export determinism and map-resolution tests pass on the full fixture corpus and the emitted markdown is byte-identical to the reference impl on redline-free constructs.

**M2 — Validator + ratification.** Gates G1–G5, diff rendering, per-edit accept/reject. ✅ when the fuzz suite catches 100% of injected corruptions and a hand-run with DHSChat produces a passing validation on a real RIA excerpt.

**M3 — Injection + ZIP writer.** ✅ when §9.4 holds on every fixture: no repair prompt, correct attribution, accept-all equivalence, comments threaded.

**M4 — Run Review end-to-end.** Prompt packaging, persona import, session files, chunk mode. ✅ when a full review of a real document completes clipboard-only and the audit sidecar is complete.

**M5 — Train Persona wizard.** Incl. exemplar mining. ✅ when a persona built in the wizard round-trips (import → prompt assembly) and saves cleanly into StickShift.

**M6 — Respond to Review.** ✅ when replies land as threaded comments visible in Word with correct parents and the recommendations render as point comments.

Ship after M4; M5/M6 follow.

---

## 16. Known risks & mitigations

| Risk | Mitigation |
|---|---|
| Word repair prompt | Never reformat XML; copy untouched parts byte-for-byte; complete comment part linkage; no-op round-trip test in CI |
| LLM paraphrases untouched text | G2 byte-equality gate + repair-prompt button; exemplars in prompt showing full-document echo |
| Curly braces mangled in chat UI | Fenced-block envelope both directions; parser takes last fence |
| Token limits on long docs | Chunk-by-heading mode with per-chunk G2 |
| Emphasis/`**` boundary edits | Synthetic-char snapping (§5.2); reject synthetic-only edits |
| Paragraph-mark tracking subtleties | Dedicated fixtures for whole-para insert/delete incl. list items and last-paragraph edge case (body's trailing `sectPr` paragraph must never be deleted) |
| `file://` clipboard quirks | `execCommand` fallback retained |

---

## 17. Open items for the maintainer (Luke)

1. Confirm the default author string (`AutoReviewer — {persona}`) satisfies the attribution requirement for documents that leave the branch (Jim's terminal review).
2. Test once, early: does DHSChat's UI preserve `{++ ++}` inside fenced blocks verbatim on copy-out? (Expected yes; if not, switch the envelope to a downloadable/uploadable `.md` attachment flow.)
3. Decide whether Respond-flow recommendation comments should carry a machine-parseable prefix (e.g., `[AR:accept]`) to enable a future v2 auto-apply pass.

---

## 18. Milestone 6b - Dated Rulings (2026-07-15)

### 18.1 Respond Flow Grammar & Coverage Validation
The Respond to Review flow accepts a structured response block where comments (`[Cn]`) and revisions (`[Rn]`) are addressed via a strict grammar:
- Every comment label `[Cn]` must receive a reply wrapped in `{>>...<<}`.
- Revisions `[Rn]` must receive a recommendation with a decision prefix: `[AR:accept]`, `[AR:reject]`, or `[AR:discuss]`.
- Comment replies may optionally request resolution using the `[AR:resolve]` prefix.
- The validator enforces that every annotation label present in the document is addressed exactly once. Any missing, duplicate, or unknown labels will fail the G1/G2 validation gates.
