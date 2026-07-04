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

export function buildValidResponse(rng, exportedMarkdown, sourceMap) {
  const runs = allRuns(sourceMap);
  if (!runs.length) return exportedMarkdown;
  const k = Math.min(1 + randInt(rng, 4), runs.length);
  const chosenRuns = shuffled(rng, runs).slice(0, k);
  const edits = chosenRuns.map((run) => randomEditForRun(rng, exportedMarkdown, run));
  // Descending by mdStart so each splice only affects text to its own right, already
  // spliced text stays untouched -- but two edits from *adjacent* runs can share the exact
  // same boundary position (e.g. a zero-width ins right where the next run's comment
  // span begins). Break ties by mdEnd descending too, so the wider span at that position
  // is spliced first, leaving the narrower/zero-width one's boundary undisturbed for its
  // own splice right after.
  edits.sort((a, b) => b.mdStart - a.mdStart || b.mdEnd - a.mdEnd);
  let response = exportedMarkdown;
  for (const edit of edits) {
    response = response.slice(0, edit.mdStart) + edit.raw + response.slice(edit.mdEnd);
  }
  return response;
}
