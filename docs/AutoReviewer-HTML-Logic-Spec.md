# AutoReviewer (HTML) — Distilled Logic Spec

First-principles logic distillation — 2026-06-22. **Language-agnostic.** This is
the *what and why* extracted from the VBA build, reframed for the sandbox. It is
**not** the JavaScript spec — that is the next artifact, derived from this one
after you've ruled on the open decisions in Part VI.

**How to read this.** Part I is the durable logic the HTML tool must preserve
(the contracts that make AutoReviewer *AutoReviewer*). Part II is the loop,
reframed for the sandbox. Part III is the host-tax that the VBA carried and the
sandbox *drops* — read this as the alignment check: if anything here is logic
you considered essential rather than contortion, flag it. Part IV is the honest
list of what gets *harder* in OOXML. Part V fixes the scope boundary. Part VI is
what you decide before the JS spec is written.

---

## §0 — The reframe (one move)

The VBA AutoReviewer has **two host touches**, and almost everything that makes
it complicated is the cost of those two touches:

1. **Export / anchor** — open the user's `.docx` over COM, stamp anchors,
   extract text + comments + revisions into a markdown payload.
2. **Apply** — read the model's JSONL, open Word, write the edits back as
   tracked changes, strip anchors, leave a clean reviewed doc.

In the sandbox, **the user uploads the `.docx` into the browser.** That single
fact collapses both touches into two passes of *one offline HTML tool* over an
in-memory copy:

```
  [user drops draft.docx into the tool]
        │
   PASS 1 "Prepare"  ── parse + anchor + extract ──▶ markdown payload (to clipboard)
        │                                                    │
        │                                          [DHSChat: hot co-thinker → ratify → cold serializer]
        │                                                    │
   PASS 2 "Apply"  ◀── paste JSONL back into the same open tool ──┘
        │
        ├──▶ reviewed.docx  (tracked changes baked into the XML — download)
        └──▶ markdown receipt (counts + unaddressed comments + fingerprints — copy back to log)
```

The draft never leaves the browser. The tool never touches the host. The user
places every file. This is rung-2, data-class, offline — and it is the *same*
deterministic core the VBA already has, minus the COM.

**Boundary that does not move:** the tool is the deterministic substrate
(anchor, extract, validate, apply). The **judgment** still lives in the DHSChat
assistants (hot co-thinker, cold serializer, ratification between them). Those
SKILL.md templates are unchanged by this refactor. We are porting the
*plumbing*, not the *reviewer*.

---

## Part I — The durable logic (must survive the port)

### I.1 The no-fabrication backbone: the anchor model

Every proposed edit must cite a location that provably exists in the source.
That is the property the whole tool is built to guarantee, and it is enforced by
**anchors**: the tool assigns a stable id to every addressable element, the
model may only reference those ids, and an edit citing an unknown id is refused
on apply.

The anchor namespace (carried over verbatim — these prefixes *are* the contract
the assistants are trained against):

| Anchor | Over what | Numbering |
|---|---|---|
| `AR_PARA_NNNNN` | each paragraph | document order, 1-based, zero-padded to 5 |
| `AR_CELL_NNNNN` | each table cell | document order |
| `AR_FN_NNNNN` | each footnote | document order |
| `AR_COMMENT_N` | each existing comment | index into the comments collection, 1-based |
| `AR_REV_NNNNN` | each existing tracked revision | index into the revisions collection, 1-based |

**The invariant that makes anchoring free:** the numbering is a *pure function
of the unedited document*. The VBA relies on this so that the export pass and a
later apply pass derive identical ids without persisting anything. In the
sandbox you get this for free *within a session* (the parsed doc is held in
memory across both passes), and you get it *across sessions* as long as the
derivation is deterministic over the same input file. **Keep the derivation
deterministic and document-order-based** so a re-open of the same `.docx`
re-anchors identically.

### I.2 The seven change types (semantics)

The model's only output vocabulary. Each is keyed to an anchor.

| change_type | Means | Target anchor | OOXML it becomes |
|---|---|---|---|
| `replace_text` | substitute a span | range anchor | `<w:del>`(old) + `<w:ins>`(new), **minimal span only** |
| `delete_element` | remove the element | range anchor | wrap content in `<w:del>` (+ mark para mark deleted if whole para) |
| `add_comment_only` | attach a fresh comment | range anchor | new comment in `comments.xml` + reference in body |
| `reply_to_comment` | reply to an existing comment | `AR_COMMENT_n` | threaded reply (`commentsExtended` parent link) |
| `accept_revision` | accept existing tracked change(s) in range | range anchor | resolve existing `<w:ins>`/`<w:del>` to final |
| `reject_revision` | reject existing tracked change(s) in range | range anchor | resolve existing `<w:ins>`/`<w:del>` to original |
| `add_footnote` | attach a citation | range anchor (after `old_text` if given) | footnote in `footnotes.xml` + tracked-inserted callout |

### I.3 The JSONL line contract (frozen)

The serializer emits one JSON object per line. Fields:

- `bookmark_id` — **required, string.** The anchor.
- `change_type` — **required, string.** One of the seven.
- `old_text` — optional. Scopes/locates the span inside the anchor.
- `new_text` — the replacement / footnote body.
- `add_comment` — comment or reply text.
- `apply_change` — bare boolean (a quoted `"true"` is ignored).
- `confidence` — optional, normalized.

This is already specified bit-exactly in `ref/jsonl_contract.py`, and that file
is the source of truth. **In JS it becomes the implementation directly — no
twin.** Behaviors that must port *exactly* (they are deliberate, not accidental):

- Single left-to-right tokenizer (not regex / not whole-line `indexOf`), so a
  key name appearing inside another field's *value* is never mistaken for a key.
- A closing quote is recognized only after an **even** run of backslashes
  (escape-parity correctness).
- Duplicate keys → **first occurrence wins.**
- Whitespace trim covers space, tab, CR, LF (not just spaces).
- `bookmark_id` and `change_type` must be present *as strings* or the line fails
  to parse (`PARSE`).
- The JSON unescape set is exactly `\` `"` `/` `b` `f` `n` `r` `t` and `\uXXXX`
  (with surrogate-pair combining); any other escape drops the backslash and
  keeps the char.

### I.4 The validation contract (frozen — check order is part of the contract)

After parse, each line is validated to a stable reason code. **Order matters**
(first failing check wins); mirror `ref/jsonl_contract.py:validate_change`:

1. empty `bookmark_id` → `MISSING_BOOKMARK`
2. empty `change_type` → `MISSING_CHANGE_TYPE`
3. unknown type → `UNKNOWN_CHANGE_TYPE`
4. `replace_text` with empty `new_text` → `REPLACE_REQUIRES_NEW_TEXT`
5. `add_comment_only` with empty `add_comment` → `COMMENT_REQUIRES_TEXT`
6. `reply_to_comment` not targeting `AR_COMMENT_` → `REPLY_REQUIRES_COMMENT_TARGET`; empty text → `REPLY_REQUIRES_TEXT`
7. `accept/reject_revision` targeting a comment → `REVISION_REQUIRES_RANGE_TARGET`
8. `add_footnote` targeting a comment → `FOOTNOTE_REQUIRES_RANGE_TARGET`; empty `new_text` → `FOOTNOTE_REQUIRES_TEXT`

The golden vectors in `tests/vectors/` exercise these codes. They transfer to JS
as-is and become your acceptance suite for the parser/validator — the one place
where keeping the VBA's exact algorithm buys you a free, pre-written test corpus.

### I.5 The session-binding gate (frozen — default-deny)

Anchor ids are generic ordinals, so JSONL from *document A* can apply cleanly to
*document B*. The gate prevents this. The serializer's **first line is a meta
line**:

```json
{"meta": "autoreviewer", "session": "<token>", "count": N}
```

On apply, before producing any output: verify the `session` token equals the
token issued at Prepare, and `count` equals the number of edit lines that
follow. **Any mismatch refuses the whole payload** — no partial apply. (Codes:
`NO_EXPORT_TOKEN`, `NO_PAYLOAD`, `META_MISSING`, `TOKEN_MISMATCH`,
`COUNT_MISMATCH`; see `ref/session.py`.)

Two supporting rules port with it:

- **Fence filtering.** The serializer wraps its lines in one ```` ```jsonl ````
  block with prose notes after it. Take only the lines between the first fence
  and the next; if no fence, take all non-blank lines. Both paths gate
  identically. (`ref/session.py:filter_payload_lines`.)
- **The token is the payload fingerprint.** A stable hash over the Prepare
  payload, issued at Prepare and echoed by the serializer. *(Algorithm choice =
  open decision VI.2.)*

### I.6 The comment-coverage warn-gate (frozen)

A silently unanswered reviewer comment is the worst failure mode — a false
negative the tool exists to catch. At Prepare, record every `AR_COMMENT_` id. At
Apply, compute which received **no** `reply_to_comment` *or* `add_comment_only`
line (only those two count as "addressing"; a `replace_text` that happens to hit
a comment does not). If any remain unaddressed, **warn before producing the
file** (Proceed / Abort) and list them. A no-action ruling is legitimate — but
it must be a visible choice, never an omission. (`ref/coverage.py`.)

### I.7 Minimal-diff surgery

`replace_text` tracks **only the differing span**, not the whole element: change
one word and the revision is one word, not a paragraph delete-and-reinsert. The
VBA does this by narrowing to `old_text` (if given), then diffing to the minimal
changed middle. **This is the single highest-value behavior to preserve** — it
is what makes the output read like a human reviewer's redline rather than a
machine's. In OOXML this is the run-splitting problem (Part IV.1).

### I.8 Normalization + token hygiene

- **Punctuation normalization.** Match `old_text` against the doc on a
  normalized form (smart quotes/dashes → ASCII), so the model's straight quote
  finds the doc's curly one. Normalization must be **1:1** so positions map back
  onto real offsets.
- **Anchor-token stripping.** No `AR_*` token may ever land in written text
  (insertions, comments, footnotes). Strip before writing.
- **Authoring.** All AI insertions, deletions, comments, and footnotes are
  authored **"AutoReviewer"** with a fixed date. *(In OOXML you set `w:author`
  and `w:date` directly — see III.3, this is where the sandbox is strictly
  better than VBA.)*

### I.9 The hot / cold / ratify protocol (unchanged — DHSChat-side)

For completeness, since the tool's payload feeds it: the review leg is **hot
co-thinker** (3-turn: themes → ratified blocks → final packet) → **human
ratifies** → **cold serializer** (turns the ratified packet into strict JSONL,
never re-decides). The HTML tool sits *underneath* this and is unaffected by it,
except that Prepare produces the payload the hot assistant reads and Apply
consumes the cold assistant's JSONL.

### I.10 The Prepare payload (the markdown bus)

What Prepare emits for the model to reason over — section markers carried over
so the assistants' training still matches:

- `<<PAYLOAD_FINGERPRINT: …>>` — the session token.
- `<<DOCUMENT_TEXT_START>> … <<DOCUMENT_TEXT_END>>` — the readable text.
- `<<BOOKMARK_INDEX_START>> … <<END>>` — one line per anchor:
  `AR_PARA_00037 | type=paragraph | "first 200 chars…"`.
- `<<FOOTNOTES_START>> … <<END>>`.
- `<<COMMENTS_START>> … <<END>>` — `## AR_COMMENT_n`, Author, Date, scope
  sentence, body.
- `<<REVISIONS_START>> … <<END>>` — when the doc already has tracked changes.

---

## Part II — The loop, reframed for the sandbox

### II.1 One tool, two passes, one in-memory session

The tool holds the parsed `.docx` in memory for the life of the browser tab.

**Prepare pass.** Input: the uploaded `.docx`. The tool unzips, parses
`document.xml` / `comments.xml` / `footnotes.xml`, derives anchors (I.1),
extracts the payload (I.10), computes the session token, and copies the payload
to the clipboard. Output shape: **markdown the model reasons with** (OKF
`markdown-continue`). Nothing is stored; the model just reads it.

**Apply pass.** Input: the JSONL pasted back (same tab, doc still loaded).
Filter fences → check session gate (I.5) → parse + validate every line (I.3–I.4)
→ compute coverage (I.6) → if clean (or operator proceeds past the warn-gate),
write each edit into the in-memory XML as tracked changes (Part I.2 mapping) →
re-zip → trigger download. Output shape: **file + receipt** (OKF `file-receipt`,
see II.2).

### II.2 The end shape: `file-receipt`, not bare `file-terminal`

The VBA's Trace/Log sheets were "the defensible artifact — the audit lineage is
the product." A sandboxed tool can't write a log; it has no host. So preserve
the lineage the OKF-sanctioned way: **Apply emits a short markdown receipt** the
operator carries back to the knowledge base —

```
reviewed: draft.docx | session 4A1F…9C  | 14 applied, 2 skipped
unaddressed comments: (none)
skipped: AR_PARA_00112 (old_text not found); AR_PARA_00031 (new_text == existing)
```

The reviewed `.docx` is the deliverable (downloaded, lands wherever the human
saves it, does **not** re-enter the bundle); the receipt is the trace (stored).
That split is exactly the OKF rule "the bundle holds knowledge; tools produce
deliverables," and naming the receipt is how "every action is auditable"
survives the move off VBA.

### II.3 No working copy needed

The upload *is* a copy. The source of record on the user's disk is never opened,
never mutated — the sandbox cannot reach it. The VBA's whole `*_AR`
working-copy dance (II in the old README) exists only because VBA opened the
real file; it disappears here for free.

---

## Part III — What the sandbox *drops* (the alignment check)

Read each of these and confirm it was host-tax, not vision. If any is logic you
intended, flag it before the JS spec.

**III.1 The two-pass apply ordering.** The VBA applies text/comment edits first,
then `accept/reject_revision` second, because in a live Word session accepting a
revision can delete the very range another edit targets. In XML you are not
mutating a live document under your own feet — you compute the final markup from
a stable parse. You still need a coherent *order of reasoning* about overlapping
edits, but the COM-specific "Object has been deleted" hazard that forced the two
passes is gone.

**III.2 Pagination / ScreenUpdating suspension, O(n²) collection access.** Pure
COM-performance contortions. None of it exists in a DOM/XML parse.

**III.3 The author-identity wart (the big one).** The README's most awkward
caveat: AI insertions take their author from `Application.UserName`, *unless*
Word is signed into an M365/DHS account, in which case the account name wins and
the user must toggle an obscure Word Option. **This entire problem vanishes.** In
OOXML you write `w:author="AutoReviewer"` and `w:date="…"` directly into each
`<w:ins>`/`<w:del>`; the author is whatever the XML says, deterministically, on
every machine. This is a case where the sandbox port is not just equivalent but
*strictly cleaner* than the original.

**III.4 URL launching, persona-URL storage, Config/Personas sheets.** Workbook
chassis. The tool no longer launches DHSChat or stores assistant URLs — the
human drives the clipboard between the tool and whatever assistant they choose.

**III.5 The Python twin suite (as a maintenance burden).** The twins exist
because VBA can't be tested where it runs. In JS the tool *is* the test target —
the runtime and the tests are one language in one place. The twins don't
disappear as *tests*; they convert into the JS unit tests. What disappears is
the *twin-drift class of bug* and the obligation to write every rule twice. (Per
your own P5: a falling twin-count is the architecture getting healthier — this
takes it to zero.)

---

## Part IV — What gets *harder* in OOXML (honest risk list)

The flip side of III. These are where Word's object model was doing real work
you now own. Rank these for the JS spec; they are the build risk.

**IV.1 Run-splitting for minimal-diff (`replace_text`).** A paragraph's text is
spread across `<w:r>` runs carrying formatting (`<w:rPr>`). To track only the
differing middle, you must locate `old_text` across run boundaries, split the
runs at the change edges, wrap the deleted slice in `<w:del>`/`<w:delText>` and
the new slice in `<w:ins>`, and **copy the original run's `<w:rPr>` onto the new
runs** so bold/font survive. This is the single hardest piece and the core of
the value (I.7). Word did this for you; now the tool must. **Prototype this
first** — it de-risks the whole project.

**IV.2 accept/reject of *existing* revisions.** The doc arrives with tracked
changes already in it (`AR_REV_*`). Accepting one means materializing an
existing `<w:ins>` (unwrap, keep text) or `<w:del>` (remove text); rejecting is
the inverse. Manipulating *existing* revision markup in raw XML is fiddlier than
adding new markup. Consider whether a library handles this or whether it's
hand-rolled.

**IV.3 Comment replies (`reply_to_comment`).** Threaded replies aren't just a
second comment — modern Word links them via `commentsExtended.xml`
(`w15:commentEx` with a `paraIdParent`). Plain `comments.xml` gets you fresh
comments (`add_comment_only`) easily; *replies* need the extended part. Confirm
the target library emits it, or scope replies carefully.

**IV.4 `old_text` location across runs.** I.8's punctuation-normalized substring
match was simple in VBA because `Range.Text` is a flat string. In XML the text
is fragmented across runs; you'll reconstruct a flat normalized string with an
offset map back into the run structure to find the span. Related to IV.1.

---

## Part V — Scope boundary (V1)

**In scope: Review mode only** — the per-document routine path (Prepare → Apply).
This is the OKF "build-first" item and the one with a paying audience.

**Parked: Training mode.** Mining N redlined docs into a persona is a
read-many-files host operation that produces a SKILL.md, run rarely. It is a
*separate tool* (or stays VBA) and is explicitly **not** in this HTML port.
Confirm you agree — it's the biggest scope call here.

**Unchanged: the DHSChat assistants.** Co-thinker, serializer, incorporator,
researcher, citation templates — all stay as-is. This refactor touches the
deterministic substrate only.

**Out: everything the V1 anti-scope list already excludes** — auto-apply without
tracked changes, automation bridges, auto-detecting authors, non-`.docx` inputs.

---

## Part VI — Open decisions (rule on these before the JS spec)

1. **One tool or two?** One tool with Prepare/Apply modes sharing an in-memory
   doc (proposed) — or two separate tools (Prepare emits an anchored artifact,
   Apply re-ingests doc + JSONL)? One-tool is simpler for the operator and makes
   the session token trivially consistent; two-tool is more composable but
   re-derives anchors on re-ingest (fine, given I.1's determinism). *Leaning
   one-tool.*

2. **Session-token algorithm.** Reuse the exact VBA `ArContentFingerprint`
   (two 31-bit polynomial lanes) so `tests/vectors/fingerprint_vectors.txt`
   transfers as a free JS test — *or* use any stable hash (e.g. a simple
   content hash) since the token is now self-contained within one tool and
   never has to match a VBA-produced token? *Leaning reuse, purely for the free
   vectors.*

3. **OOXML library vs. hand-rolled.** A zero-dependency generator (e.g. the
   `docx` library) is clean for *new* markup but weaker at *editing an existing
   doc's* runs/revisions (IV.1, IV.2). A browser OOXML editor core handles
   existing-doc editing but is heavier and must be vendored offline (no CDN, per
   P3). Decide per the IV risks. This is the load-bearing technical choice.

4. **Receipt format.** Fix the exact receipt schema (II.2) now — it's the audit
   contract, and you'll want it stable across runs for the log.

5. **Input ergonomics.** Prepare takes a file upload (sandboxed picker). Does
   Apply re-use the in-memory doc from Prepare (one-tool) or also take a fresh
   upload (two-tool)? Tied to VI.1.

6. **Footnote/comment XML parts.** Confirm the target documents actually use the
   parts you'll write (`footnotes.xml`, `commentsExtended.xml`) — TSA templates
   may differ, and a missing part means creating it from scratch.

---

*Next artifact (after your edits): the JavaScript spec — the same logic bound to
concrete OOXML operations, library choices, module layout, and the JS test
suite seeded from `tests/vectors/`. Then code from the JS spec.*
