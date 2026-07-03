# M3b — Injection (ooxml/inject.js, ooxml/comments.js, ooxml/serialize.js)

**Status: APPROVED — 2026-07-03.** This is the authoritative implementation plan.
Rulings on all design decisions are FINAL and incorporated inline below; no open
questions remain. Implementation happens in a fresh session that reads this file.

## Implementation protocol (binding)

- Branch: `claude/m3b-injection`. One PR, M3b only.
- Test-first, in this order: write each test file from the Tests section and see it
  red BEFORE implementing the code it covers.
- Commit after each test file goes green — not one batch at the end. The operator may
  lose quota mid-milestone and will resume from the repo in a fresh session; every
  green-test commit is a valid resume point.
- Never relax a test or gate to pass. Golden files and fixtures are read-only.
- Scope guard: commentsIds.xml and commentsExtensible.xml are OUT of scope (ruling
  D6 below). Do not implement them in this milestone.

## Rulings summary

- **D1 — approved as written.** Whole-paragraph detection in criticmarkup/parse.js.
- **D2 — approved as written.** resolvePoint returns a discriminated result.
- **D3 — approved as written.** Whole-paragraph delete bypasses snapBoundary; G4
  locked-range check still runs unconditionally.
- **D4 — approved as deferred**, with one requirement: the thrown error message must
  tell the USER the workaround — "This paragraph contains existing tracked changes.
  Accept or reject them in Word first, then re-export and re-run the review."
- **D5 — approved as written.** Best-effort list detection; never fabricate a numId
  with no numbering definition behind it.
- **D6 — smaller cut.** M3b creates/extends comments.xml, commentsExtended.xml, the
  document rels entries, and [Content_Types].xml overrides ONLY. commentsIds.xml and
  commentsExtensible.xml are deferred to a follow-up implemented against a
  Word-authored fixture (fixtures/comments-word-authored.docx) that the operator will
  provide. upsertComments's signature may reserve keys for the deferred parts, but no
  code writes them in M3b.

---

## Context

M3a (ZIP writer, no-op round trip only) is merged and Word-verified. M3b is the actual
injection milestone: turn accepted CriticMarkup edits into native Word tracked-change
XML in document.xml, create/extend the comment parts, and wire ratification → inject →
writeZip → download. This is the riskiest milestone (direct OOXML surgery on
document.xml's DOM, where a schema-ordering mistake is exactly what produces Word's
repair prompt), so this plan states the run-splitting algorithm and edit-ordering
strategy precisely enough to check against spec §9.1.

Two real design gaps turned up during research that spec §9.1's prose doesn't spell
out in enough detail to code directly, and both needed to be resolved before an
algorithm could be written down:

- resolvePoint (M2) only returns {bodyPath} — no run/char granularity at all, even
  for an ordinary "insert one word mid-sentence" point edit. It also flagged
  whole-paragraph insert as explicitly unresolvable, deferred to M3.
- Position alone can't distinguish "insert text into this paragraph" from "insert a
  whole new paragraph here" — both can legally resolve to the exact same markdown
  offset (e.g. inserting at the very start of a no-prefix paragraph's first run ==
  inserting a new paragraph immediately before it). These are structurally different
  DOM operations (new run inside paragraph N vs. new sibling <w:p>), so geometry alone
  is insufficient; the token's raw shape in the response (alone on its own line, per
  spec §4's literal definition) is the actual discriminator.

## Design decisions (RULED — final; see Rulings summary)

**D1 — Whole-paragraph detection lives in criticmarkup/parse.js, not the source map.**
For each ins/del token, compute a wholeParagraph: true/false flag from the raw
response text: the token is whole-paragraph if everything from the start of its line
to rawStart is blank (back to the nearest "\n\n" or start of string) and everything
from rawEnd to the end of its line is blank (forward to the nearest "\n\n" or end of
string). This is spec §4's own definition ("a line consisting solely of one token")
applied literally at the raw-text level, and it's a required signal — see the "why not
position-based" note in the algorithm below. Only ins/del get this flag; spec doesn't
define whole-paragraph substitution/comment, so those aren't special-cased (a
whole-line {~~old~>new~~} is treated as an ordinary substitution over the paragraph's
real-text content — synthetic prefix snapped away as usual).

**D2 — resolvePoint's return shape changes from {bodyPath} to a discriminated
result:** {kind: "run", bodyPath, runIndex, charOffset} (ordinary in-run point) or
{kind: "paragraphBoundary", bodyPath, edge: "before"|"after"} (whole-paragraph point,
matched against an exact block boundary). Nothing outside validate.js consumes the old
shape yet (ratify.js never reads .anchor), so this is a clean extension, not a
breaking change to any shipped behavior — existing sourcemap.test.js/validate.test.js
cases get updated accordingly.

**D3 — Whole-paragraph delete bypasses snapBoundary+resolveRange entirely.** For a
prefixed block (heading/list), snapBoundary would snap the edit's start forward past
the synthetic "# "/"- " prefix, silently downgrading a whole-paragraph delete into an
ordinary content-only delete and losing the paragraph-mark deletion — which is the
entire point of the operation (Word merges paragraphs on accept only if the mark
itself is flagged). So: when wholeParagraph: true and the edit's raw [mdStart, mdEnd)
exactly equals some block's full [block.mdStart, block.mdEnd) (prefix included),
resolve directly to {bodyPath, wholeParagraph: true} — skip the run-coverage check.
G4's locked-range check still runs unconditionally against the original span either
way.

**D4 — Whole-paragraph delete scope: plain-text runs and hyperlinks only.** A
paragraph containing locked content (image/field/content-control) can never reach this
path anyway — its span would overlap a locked range and G4 rejects it upstream. A
paragraph containing pre-existing w:ins/w:del being whole-paragraph-deleted is a real
edge case (Word's delete-of-an-insertion semantics are non-trivial) and is DEFERRED by
ruling: inject.js throws a clear, named error. Per the ruling, the error message must
tell the user the workaround: "This paragraph contains existing tracked changes.
Accept or reject them in Word first, then re-export and re-run the review."

**D5 — New-paragraph list detection is best-effort.** A whole-paragraph insert's text
may itself contain the exporter's own bullet prefix (spec §4's example:
{++- new bullet++}). inject.js strips a recognized leading "- "/"  - "/"1. " pattern
from the inserted text and, only if the new paragraph is being inserted immediately
adjacent to an existing list paragraph, copies that neighbor's w:numPr onto the new
paragraph so it's a real Word list item. If there's no adjacent list to copy from, the
prefix is left as literal text (renders as a paragraph starting with "- ", not a true
Word bullet) rather than fabricating a numId with no numbering definition behind it.

**D6 — Comment parts scope (RULED: smaller cut).** M3b implements comments.xml,
commentsExtended.xml (threading/done), the document rels entries, and
[Content_Types].xml overrides. commentsIds.xml and commentsExtensible.xml are NOT
implemented in M3b: they are a follow-up milestone implemented against a Word-authored
fixture (fixtures/comments-word-authored.docx, operator-provided) so the schemas are
copied from XML Word itself wrote rather than reconstructed from memory. This is the
low-confidence/high-cost zone (repair-prompt-critical), so no schema guessing here.

## The run-location + run-splitting algorithm (per spec §9.1)

### Locating the target w:p from bodyPath

Replicate walkBodyTracked's exact traversal: walk body.children, incrementing a
counter for every child regardless of type (matching export.js's unconditional
bodyIdx++). When the counter matches bodyPath[0]: if bodyPath.length === 1, that child
is the target w:p. If bodyPath.length === 4 (table cell), that child must be the
w:tbl; descend kids(tbl,'tr')[bodyPath[1]] → kids(tr,'tc')[bodyPath[2]] →
kids(tc,'p')[bodyPath[3]].

### Locating the target w:r from runIndex

Replicate buildSegments's exact counting order (this is the part that must match
export.js byte-for-byte or indices desync): walk the paragraph's children; every w:r
that is a direct child of p increments the counter and is a candidate return value;
every w:r nested inside w:ins/w:del/w:hyperlink/w:smartTag/w:fldSimple/w:sdt also
increments the counter (via the same recursive walk collectRunsTextIndexed does) but
is never itself a target, since sourceMap.blocks[].runs only ever records direct-child
"text" segs. In practice this means: run a single recursive counter matching
buildSegments's traversal, and only ever return a hit when the count lands on a direct
child w:r of p.

### Why right-to-left works, and why it's a per-run "peel from the right", not a per-edit rule

G1 (no nesting/block-crossing) + G2 (byte-exact strip) together guarantee something
useful: since tokens never nest and appear in the response in strict document order,
and strip maps each token to a determinate, non-overlapping slice of the stripped text
in that same order, two edits' resolved spans can never overlap. So overlap-detection
isn't a new gate — it's a structural invariant already enforced by G1+G2, asserted
defensively in inject.js (cheap insurance, not a new validation pass).

Multiple edits can target the same run, though (the adjacent-edits case — e.g. two
insertions in the same sentence, or a delete immediately followed by an insert). This
is where ordering is load-bearing, and it's per run, not globally:

- For a given run, collect all edits touching it (each edit may touch several runs if
  its span crosses run boundaries — see below) and process them right to left by
  charOffset within that run (rightmost/highest offset first).
- The first (rightmost) edit touching a run splits the run's original text at
  [charStart, charEnd): a left remainder piece [0, charStart) (kept as a plain clone,
  becomes the new "current" node for that run), the edit's own wrapped piece
  (<w:del>/<w:ins>/del-then-ins for sub), and a right remainder piece [charEnd, len)
  (plain clone, inserted after the wrapped piece and never touched again). The
  original run node is replaced in place by this 1-3 node sequence (insertBefore +
  removeChild — not relying on replaceWith, since it isn't guaranteed present on the
  injected DOM implementations @xmldom/xmldom/happy-dom use).
- The second (further-left) edit on the same run operates on the left remainder node
  left behind by the first — its own charStart/charEnd (still expressed in the run's
  original coordinate space, never renumbered) are guaranteed <= the first edit's
  charStart by the no-overlap invariant above, so they index correctly into that
  remainder's text, which is exactly original[0:firstEdit.charStart). This repeats for
  however many edits land on one run — each one only ever trims further off the right
  end of whatever's left, which is exactly why right-to-left is the correct (and only
  non-error-prone) order: offsets computed once, up front, against the pristine
  original text stay valid throughout, with no coordinate translation needed at any
  step.
- A run with zero edits touching it is never visited at all — untouched siblings,
  including pre-existing w:ins/w:del, comment markers, fields, images, are left
  completely alone.

Multi-run edits (a del/sub/anchored-comment span crossing a run boundary):
resolveRange already returns one triple per touched run, with per-run-local
charStart/charEnd — for the first run touched, charEnd is clamped to that run's own
end (Math.min(run.mdEnd, mdEnd)); for interior runs, the whole run is covered
(charStart=0, full length); for the last run, charStart=0 and charEnd is the local end
point. Each triple is treated as an independent per-run sub-operation using the exact
same peel-from-the-right mechanic above — no special-casing needed, it falls out of
resolveRange's existing math. Each physical <w:del>/<w:ins> element created gets its
own fresh w:id (matching how Word itself represents a multi-run change — there's no
requirement that a logically-one edit share an id across runs, only that each id be
document-unique).

Global processing order: group accepted+resolved edits by bodyPath (paragraphs never
interact structurally with each other, so order across paragraphs doesn't matter).
Within a paragraph, process right-to-left by position — in practice this falls out
automatically by processing runs from highest runIndex to lowest, and within a run, by
the per-run peel described above. Whole-paragraph insert/delete and comment-marker
placement (see below) are interleaved into this same right-to-left pass by anchor
position, so there's one ordering rule for the whole paragraph, not several.

Revision w:id allocation: single counter, seeded at max(existing w:ins/w:del w:id)+1
scanned from the pre-mutation document, incremented once per <w:ins>/<w:del> element
created. Allocated in the same right-to-left pass purely because that's the natural
order to write the code in — Word doesn't require any particular numeric ordering,
only uniqueness, so this isn't load-bearing the way the run-splitting order is.

### xml:space="preserve"

Applied uniformly: whenever a <w:t> or <w:delText> is created or has its text content
set/split (original run pieces, new inserted text, delText renames), if that specific
node's own text starts or ends with whitespace, set xml:space="preserve" on it.

### Whole-paragraph delete mechanics

Once resolved to {bodyPath, wholeParagraph: true} (D3): locate the w:p, wrap every
direct-child w:r in a fresh <w:del> (rename its w:t→w:delText; per D4, throw the
user-actionable error if a direct child is w:ins/w:del/w:fldSimple/w:sdt — locked
content can't reach here per G4, and w:hyperlink is supported: wrap each of its inner
runs in <w:del> the same way, leaving the <w:hyperlink> wrapper itself in place). Then
flag the paragraph mark: ensure pPr exists (create as w:p's first child if missing),
then ensure pPr/rPr exists — inserted at the schema-correct position: w:rPr must
appear before w:sectPr/w:pPrChange if either exists on that paragraph (this is the
concrete mechanism behind the last-paragraph/sectPr known risk — blindly appending
would produce schema-invalid XML on exactly the paragraph most likely to have a
sectPr, i.e. the last one in the body). Prepend a <w:del w:id w:author w:date/> as
rPr's first child (per CT_ParaRPr's own ordering: ins/del before the formatting
properties).

### Whole-paragraph insert mechanics

Resolved via the extended resolvePoint (D2) to {bodyPath, edge}. Build a fresh <w:p>:
copy pPr from the anchor paragraph for style continuity (heading/list style, indent)
minus sectPr/pPrChange (those never propagate to a new paragraph); set the new
paragraph mark's own pPr/rPr/ins (mirroring the delete case's position rules); strip a
leading list-prefix from the inserted text and copy numPr from the anchor only if the
anchor itself is a list item (D5); build the run(s) for the text itself, wrapped in
<w:ins>. Insert the new w:p as a sibling before/after the anchor paragraph per edge.
For multiple whole-paragraph inserts anchored at the same edge (consecutive new
paragraphs), maintain a per-anchor "last inserted" cursor and insert each subsequent
one immediately after the previous, processing the edits in their own left-to-right
document order — the one narrow, explicitly-called-out exception to right-to-left,
needed because inserting several new siblings at one point is order-sensitive in a way
that in-run splitting isn't.

### Comments (marker placement in document.xml)

Comments don't wrap/modify run content — commentRangeStart/commentRangeEnd/
commentReference are sibling elements. If a boundary falls mid-run, that run is split
using the same peel-from-the-right primitive (just inserting a marker at the split
point instead of a del/ins wrapper); if a boundary falls exactly at a run edge, no
split is needed. Point comments place both markers adjacent at the same resolved
position. Comment w:id is a separate counter from the revision-id counter (spec step
6's max(existing)+1 is its own namespace), seeded from existing
commentRangeStart/w:comment ids.

## ooxml/comments.js

upsertComments(existingParts, newComments) where existingParts is {commentsXml,
commentsExtendedXml, relsXml, contentTypesXml} (any may be null/absent; keys for the
D6-deferred parts may be reserved in the shape but are never written in M3b) and
newComments is [{id, author, date, text, parentId?}]. Returns the same shape with
mutated/created XML strings for whichever parts changed. Create-from-scratch each part
when absent (matching its part's XML declaration/namespace exactly), extend in place
when present. Generates a fresh 8-hex w14:paraId per new comment (random, checked for
uniqueness against existing paraIds in the doc — collision astronomically unlikely but
checked defensively, matching add_comment_threading's own paraId-generation spirit in
fixtures/generate.py). Threading (w15:paraIdParent) is supported in the function
signature/tests now per spec §9.2, even though M3b's Run-Review wiring never passes a
parentId (all Run-Review-created comments are top-level; replies are M6/Respond-flow
scope).

## ooxml/serialize.js

serializePart(xmlDoc, XMLSerializerImpl) → string. Thin wrapper over
XMLSerializerImpl.serializeToString(xmlDoc) (no pretty-printing to opt out of —
standard XMLSerializer doesn't reformat). Called once per mutated part before handing
the {partName: xmlString} map to writeZip.

## Wiring (src/ui/app.js's demo Run Review panel)

The current demo panel discards the raw docx bytes after calling exportDocx (keeps
only {markdown, sourceMap}). Extend it to retain docxBytes alongside exported, so
injection can independently re-unzip and re-parse document.xml (and the comment parts,
rels, content-types) from the same original bytes the source map was built from —
bodyPath/runIndex are stable, deterministic offsets into that original structure. Add
the handler for the ratification UI's existing Inject accepted edits action (already
rendered by ratify.js, currently a no-op dataset.action="inject" button): on click,
re-unzip → injectEdits(documentXmlDoc, state.acceptedEdits(), sourceMap, {author}) →
upsertComments(...) for any new comments → serializePart each mutated part →
writeZip(zip, mutatedParts) → Blob download named "{original name} — reviewed.docx".
Author string input defaults to "AutoReviewer — {persona}" (persona name is a
free-text field for now, since persona files aren't wired until M5).

## Tests (written first, red before implementation)

- **tests/inject.acceptAll.test.js** — spec §13 invariant #3, written and red before
  inject.js has a real implementation: for each clean fixture, generate random valid
  edit sets (reusing the existing property-suite generator from
  tests/validate.property.test.js where possible), run them through validate →
  injectEdits → a small programmatic accept-all transform (unwrap w:ins, remove w:del
  including its content, drop nothing else) → extract text, and assert it equals
  applyEdits-style reconstruction of what the model's response actually proposed (i.e.
  strip⁻¹: apply ins/del/sub to the exported markdown the same way the response did).
  This is the primary correctness net for the whole milestone.
- **tests/inject.wholeParagraph.test.js** — dedicated cases for: whole-paragraph
  insert (plain paragraph, list item, at document start, at document end, two
  consecutive inserts at the same anchor), whole-paragraph delete (plain paragraph,
  list item, the actual last body paragraph carrying sectPr in its own pPr — asserting
  rPr lands before sectPr in the serialized output), and the D4 scope-cut (asserting
  the clear, user-actionable thrown error, not corruption).
- **tests/inject.adjacentEdits.test.js** — two+ edits landing on the same run (two
  insertions, a delete immediately followed by an insertion, a delete then a comment
  boundary in the same run) — the case the right-to-left peel algorithm exists for.
- **tests/inject.runSplitting.test.js** — unit-level tests of the
  run-location/splitting primitive directly (not just end-to-end), including multi-run
  del/sub spans and xml:space="preserve" on whitespace-boundary splits.
- **tests/comments.test.js** — upsertComments create-from-scratch and extend-existing
  paths for the in-scope parts, threading (parentId) even though unused by the wiring
  yet, and round-tripped through the reader (buildCommentsData in export.js) to
  confirm a newly-injected comment is itself re-exportable/readable.
- **CI:** extend scripts/roundtrip-fixtures.mjs / verify_zip_roundtrip.py beyond the
  M3a no-op case — round-trip each fixture through a real injected edit set (not just
  writeZip(zip, {})), so the independent Python zipfile structural check also covers
  actual mutated output, not only untouched copies.

No new .docx fixtures are needed for M3b itself — whole-paragraph/adjacent-edits cases
are constructed as CriticMarkup responses against the existing exported markdown of
current fixtures (plain-paragraphs, headings-and-lists, tables), which is where these
constructs live. (fixtures/comments-word-authored.docx arrives with the D6 follow-up,
not this milestone.)
