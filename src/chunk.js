// Chunk mode (spec §6.4; two-phase split per architecture doc §7 / M4c). A document over
// the word-count threshold (prompt.js's CHUNK_WORD_THRESHOLD) is round-tripped one
// top-level-heading section at a time instead of as one giant copy/paste. Each chunk gets
// its own independent phase-1 pass (validateText: G1 grammar + G2 fidelity, checked purely
// against that chunk's own slice of the exported markdown); once every chunk has passed
// phase 1, its edits are translated into the full document's coordinate space and merged in
// document order for a single phase-2 pass (resolveEdits: G3 anchor resolution, G4
// protection, G5 sanity) against the FULL document's source map.
//
// All chunk-awareness lives in this file. validate.js's two phases (and sourcemap.js and
// inject.js, which validate.js's phase 2 calls into) never see a chunk boundary -- phase 1
// is handed a small, self-contained {exportedMarkdown, sourceMap} that looks exactly like a
// whole (small) document, and phase 2 is handed ordinary full-document-coordinate edits
// exactly like the single-document path already produces. Chunk knowledge dies right here,
// at the translation step.
import { validateText, resolveEdits } from "./validate.js";

// A block is a splittable top-level-heading boundary only if its own markdown text (read
// from the FULL exportedMarkdown, since a chunk-local slice isn't built yet at this point)
// opens with exactly one "#" -- export.js's kind: "heading" covers every level 1-6 (and
// Title, which is also rendered as "# "), and heading level isn't itself a source map
// field, so the literal prefix text is the only signal available without changing
// export.js's schema.
function isTopLevelHeading(block, exportedMarkdown) {
  return block.kind === "heading" && /^# (?!#)/.test(exportedMarkdown.slice(block.mdStart, block.mdEnd));
}

// Splits {exportedMarkdown, sourceMap} into chunks at each top-level heading. Chunk 0 always
// starts at offset 0 -- carrying the document header plus any preamble before the first
// top-level heading -- and each later chunk starts exactly at a top-level heading's own
// mdStart. A document with zero or one top-level heading yields exactly one chunk spanning
// the whole document: chunk mode degenerates to the single-document case rather than being
// something callers must special-case around.
//
// Each returned chunk's sourceMap has its blocks/locked/synthetic ranges rebased so mdStart/
// mdEnd count from 0 at the chunk's own start (matching a real single-document export's
// coordinate space) -- but bodyPath, runIndex, and charOffset are left untouched, since
// those already point at the real document's DOM structure/run-local text and have nothing
// to do with markdown-offset chunking.
export function splitIntoChunks(exportedMarkdown, sourceMap) {
  const blocks = sourceMap.blocks || [];
  const boundaries = [0];
  // The FIRST top-level heading found does not start its own chunk -- it (and its section)
  // stays merged into chunk 0 along with the document header/preamble that precedes it.
  // Only the SECOND and later top-level headings become new chunk boundaries; a document
  // with 0 or 1 top-level heading has nothing left to split on and yields one chunk.
  let seenFirstHeading = false;
  for (const block of blocks) {
    if (!isTopLevelHeading(block, exportedMarkdown)) continue;
    if (!seenFirstHeading) {
      seenFirstHeading = true;
      continue;
    }
    boundaries.push(block.mdStart);
  }
  boundaries.push(exportedMarkdown.length);

  const rebaseRange = (baseOffset) => ([s, e]) => [s - baseOffset, e - baseOffset];

  const chunks = [];
  for (let i = 0; i < boundaries.length - 1; i++) {
    const baseOffset = boundaries[i];
    const end = boundaries[i + 1];
    const inChunk = (s, e) => s >= baseOffset && e <= end;

    const chunkBlocks = blocks
      .filter((b) => inChunk(b.mdStart, b.mdEnd))
      .map((b) => ({
        ...b,
        mdStart: b.mdStart - baseOffset,
        mdEnd: b.mdEnd - baseOffset,
        runs: (b.runs || []).map((r) => ({ ...r, mdStart: r.mdStart - baseOffset, mdEnd: r.mdEnd - baseOffset })),
      }));

    chunks.push({
      index: i,
      baseOffset,
      exportedMarkdown: exportedMarkdown.slice(baseOffset, end),
      sourceMap: {
        docHash: sourceMap.docHash,
        blocks: chunkBlocks,
        locked: (sourceMap.locked || []).filter(([s, e]) => inChunk(s, e)).map(rebaseRange(baseOffset)),
        synthetic: (sourceMap.synthetic || []).filter(([s, e]) => inChunk(s, e)).map(rebaseRange(baseOffset)),
      },
    });
  }
  return chunks;
}

const MD_OFFSET_FIELDS = ["mdPos", "mdStart", "mdEnd"];

// Offset translation, the one place chunk knowledge is allowed to touch an edit: adds a
// chunk's baseOffset to every md-offset field (mdPos, mdStart, mdEnd) and to NONE of
// rawStart/rawEnd, which index into that chunk's own pasted-back response text -- a
// position that stays chunk-local no matter what document coordinate the edit resolves to
// (it's only ever used for diagnostics/repair-prompt pointing back at what the user pasted
// for that one chunk).
export function translateEdits(edits, baseOffset) {
  if (!baseOffset) return edits;
  return edits.map((edit) => {
    const translated = { ...edit };
    for (const field of MD_OFFSET_FIELDS) {
      if (typeof translated[field] === "number") translated[field] += baseOffset;
    }
    return translated;
  });
}

// Full two-phase chunk-mode pipeline. `chunks` is splitIntoChunks's output; `chunkResponses`
// is aligned 1:1 with it -- chunkResponses[i] is the pasted-back response for chunks[i].
// `sourceMap` here is the FULL document's source map (not any one chunk's), since phase 2
// resolves anchors against the real document.
//
// Phase 1 runs independently per chunk, in order, and stops at the first chunk that fails
// its own G1/G2 check -- so a bad paste is reported against the one chunk that produced it
// (chunkIndex on the failure), without needing every other chunk to already be pasted in.
// Only once every chunk has passed does phase 2 run, exactly once, against the merged,
// translated edit list -- identical in shape to what the single-document validate() would
// produce from one big response, per the chunked-vs-unchunked equivalence oracle
// (tests/chunk.test.js).
export function validateChunked({ chunks, chunkResponses, sourceMap, largestDeletionWords = 50 }) {
  const merged = [];
  for (let i = 0; i < chunks.length; i++) {
    const chunk = chunks[i];
    const phase1 = validateText({
      responseMarkdown: chunkResponses[i],
      exportedMarkdown: chunk.exportedMarkdown,
      sourceMap: chunk.sourceMap,
    });
    if (!phase1.ok) return { ...phase1, chunkIndex: i };
    merged.push(...translateEdits(phase1.edits, chunk.baseOffset));
  }
  return resolveEdits({ edits: merged, sourceMap, largestDeletionWords });
}
