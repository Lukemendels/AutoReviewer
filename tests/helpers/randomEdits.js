// Shared random-valid-edit-set generator, factored out of tests/validate.property.test.js
// so tests/inject.acceptAll.test.js can reuse the exact same generator (M3b plan: "reusing
// the existing property-suite generator from tests/validate.property.test.js where
// possible") rather than maintaining a second, subtly-different one. Behavior is
// unchanged from the original inline version -- this is a pure extraction.

// Clean fixtures only -- no pre-existing tracked changes/comments baked into the export
// (a response echoing pre-existing CriticMarkup verbatim raises its own G2 question,
// tracked separately and out of scope here).
export const CLEAN_FIXTURES = [
  "plain-paragraphs",
  "headings-and-lists",
  "tables",
  "hyperlinks-and-images",
  "bold-italic",
  "fields-and-content-controls",
];

export function mulberry32(seed) {
  return function () {
    seed |= 0;
    seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}
export function randInt(rng, maxExclusive) {
  return Math.floor(rng() * maxExclusive);
}
export function pick(rng, arr) {
  return arr[randInt(rng, arr.length)];
}
export function shuffled(rng, arr) {
  const a = arr.slice();
  for (let i = a.length - 1; i > 0; i--) {
    const j = randInt(rng, i + 1);
    [a[i], a[j]] = [a[j], a[i]];
  }
  return a;
}

export const WORDS = ["revised", "updated", "amended", "clarified", "expanded", "adjusted", "modernized", "streamlined"];
export function randomWord(rng) {
  return pick(rng, WORDS);
}

export function allRuns(sourceMap) {
  const out = [];
  for (const block of sourceMap.blocks) {
    for (const run of block.runs) out.push(run);
  }
  return out;
}

// One random edit anchored entirely within a single document-text run, so it can never
// touch synthetic scaffolding at the run's edges. ins is a zero-width point (mdStart ===
// mdEnd) at a random position in the run; del/sub/comment carve a random non-empty
// sub-span and use the *actual* substring as old text, so the response stays G2-valid.
export function randomEditForRun(rng, exportedMarkdown, run) {
  const len = run.mdEnd - run.mdStart;
  const spanLen = 1 + randInt(rng, len);
  const startOffset = randInt(rng, len - spanLen + 1);
  const spanStart = run.mdStart + startOffset;
  const spanEnd = spanStart + spanLen;
  const oldText = exportedMarkdown.slice(spanStart, spanEnd);
  const kind = pick(rng, ["ins", "del", "sub", "comment"]);
  if (kind === "ins") return { mdStart: spanEnd, mdEnd: spanEnd, raw: `{++${randomWord(rng)}++}` };
  if (kind === "del") return { mdStart: spanStart, mdEnd: spanEnd, raw: `{--${oldText}--}` };
  if (kind === "sub") return { mdStart: spanStart, mdEnd: spanEnd, raw: `{~~${oldText}~>${randomWord(rng)}~~}` };
  return { mdStart: spanStart, mdEnd: spanEnd, raw: `{==${oldText}==}{>>${randomWord(rng)} check.<<}` };
}

// D1 whole-paragraph insert: splices "\n" + token + "\n" over an ENTIRE existing "\n\n"
// gap between two blocks (or between the header and the first block) -- never after the
// LAST block, since the export's own trailing whitespace there is a single "\n", not a
// "\n\n" gap (a genuinely different, narrower construction); that edge is already covered
// by a dedicated hand-written case in tests/inject.wholeParagraph.test.js, so the fuzz
// generator doesn't need to reproduce it. Splicing over the WHOLE gap -- not a zero-width
// point inside it -- matters for G2: strip() only removes the token's own raw span, so
// the gap's two original newlines must already be sitting verbatim on either side of the
// token for the stripped result to reconstruct the original "\n\n" exactly (see
// validate.js's D7 comment for the underlying invariant this preserves). Returns null if
// the export has no blocks at all.
export function randomWholeParagraphInsert(rng, exportedMarkdown, sourceMap) {
  const blocks = sourceMap.blocks || [];
  if (!blocks.length) return null;
  const headerEnd = exportedMarkdown.slice(0, blocks[0].mdStart).replace(/\s+$/, "").length;
  const candidates = [[headerEnd, blocks[0].mdStart]];
  for (let i = 0; i < blocks.length - 1; i++) candidates.push([blocks[i].mdEnd, blocks[i + 1].mdStart]);
  // Not every inter-block gap is a blank-line "\n\n" paragraph separator -- adjacent table
  // cells in the same row, say, are also consecutive "blocks" in the source map, but the
  // text between them is markdown table syntax (" | "), not a paragraph gap. Splicing this
  // construction over anything else would leave stray characters strip() never removes,
  // failing G2. Only a gap that IS exactly "\n\n" is a real candidate.
  const gaps = candidates.filter(([s, e]) => exportedMarkdown.slice(s, e) === "\n\n");
  if (!gaps.length) return null;
  const [gapStart, gapEnd] = pick(rng, gaps);
  return { mdStart: gapStart, mdEnd: gapEnd, raw: `\n{++${randomWord(rng)}++}\n` };
}

// D3 whole-paragraph delete: the edit's declared span must exactly equal some block's own
// full [mdStart, mdEnd) -- prefix included -- for validate.js's resolveEditAnchor to route
// it through the wholeParagraphDelete path instead of an ordinary span delete. A block's
// own span always sits between two blank-line gaps (or the header/EOF) already, so this
// is "alone on its own line" by construction with no extra bookkeeping needed. Returns
// null if the export has no blocks at all.
export function randomWholeParagraphDelete(rng, exportedMarkdown, sourceMap) {
  const blocks = sourceMap.blocks || [];
  const locked = sourceMap.locked || [];
  // G4's locked check runs unconditionally against a delete's ORIGINAL declared span,
  // whole-paragraph or not (validate.js's resolveEditAnchor) -- a paragraph containing any
  // locked content (a field, a content control) can never be deleted wholesale. Excluding
  // those blocks up front keeps this generator producing only responses validate() will
  // actually accept, matching CLEAN_FIXTURES' own "no ingredients this category needs ->
  // null" convention elsewhere in the fuzz helpers.
  const eligible = blocks.filter((b) => !locked.some(([ls, le]) => ls < b.mdEnd && b.mdStart < le));
  if (!eligible.length) return null;
  const block = pick(rng, eligible);
  const oldText = exportedMarkdown.slice(block.mdStart, block.mdEnd);
  return { mdStart: block.mdStart, mdEnd: block.mdEnd, raw: `{--${oldText}--}`, blockRange: [block.mdStart, block.mdEnd] };
}

function spliceEdits(exportedMarkdown, edits) {
  // Descending by mdStart so each splice only affects text to its own right, already
  // spliced text stays untouched -- but two edits from *adjacent* runs can share the exact
  // same boundary position (e.g. a zero-width ins right where the next run's comment
  // span begins). Break ties by mdEnd descending too, so the wider span at that position
  // is spliced first, leaving the narrower/zero-width one's boundary undisturbed for its
  // own splice right after.
  const sorted = [...edits].sort((a, b) => b.mdStart - a.mdStart || b.mdEnd - a.mdEnd);
  let response = exportedMarkdown;
  for (const edit of sorted) {
    response = response.slice(0, edit.mdStart) + edit.raw + response.slice(edit.mdEnd);
  }
  return response;
}

export function buildValidResponse(rng, exportedMarkdown, sourceMap) {
  const blocks = sourceMap.blocks || [];

  // ~15% of draws also include one D1 whole-paragraph op (insert-into-a-gap or
  // delete-a-whole-block, chosen with equal probability) -- closes the fuzz-coverage hole
  // around resolvePoint's paragraphBoundary path and resolveEditAnchor's
  // wholeParagraphDelete path, neither of which randomEditForRun's ordinary in-run edits
  // (always confined inside a single run) can ever reach. Chosen BEFORE the ordinary run
  // edits below so a whole-paragraph DELETE's block can be excluded from the ordinary
  // draw -- the two would otherwise cover the identical md span and corrupt the splice. A
  // gap-insert never needs this: a gap sits strictly BETWEEN blocks, so it never overlaps
  // any run's own span regardless of which runs get chosen.
  let wholeParagraphEdit = null;
  let excludeBlockRange = null;
  if (blocks.length && rng() < 0.15) {
    if (rng() < 0.5) {
      wholeParagraphEdit = randomWholeParagraphInsert(rng, exportedMarkdown, sourceMap);
    } else {
      wholeParagraphEdit = randomWholeParagraphDelete(rng, exportedMarkdown, sourceMap);
      if (wholeParagraphEdit) excludeBlockRange = wholeParagraphEdit.blockRange;
    }
  }

  const eligibleRuns = allRuns(sourceMap).filter(
    (run) => !excludeBlockRange || run.mdStart < excludeBlockRange[0] || run.mdStart >= excludeBlockRange[1]
  );
  if (!eligibleRuns.length) return wholeParagraphEdit ? spliceEdits(exportedMarkdown, [wholeParagraphEdit]) : exportedMarkdown;

  const k = Math.min(1 + randInt(rng, 4), eligibleRuns.length);
  const chosenRuns = shuffled(rng, eligibleRuns).slice(0, k);
  const edits = chosenRuns.map((run) => randomEditForRun(rng, exportedMarkdown, run));
  if (wholeParagraphEdit) edits.push(wholeParagraphEdit);

  return spliceEdits(exportedMarkdown, edits);
}
