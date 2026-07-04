// Source map schema + offset resolution (spec §5.2): given an md offset range, return
// {bodyPath, runIndex, charStart, charEnd} triples, or throw a typed error naming the
// overlap kind (synthetic/locked) for the validator (M2) to surface.

export class SourceMapError extends Error {
  constructor(kind, message, range) {
    super(message);
    this.name = "SourceMapError";
    this.kind = kind; // "locked" | "synthetic"
    this.range = range;
  }
}

function rangesOverlap(aStart, aEnd, bStart, bEnd) {
  return aStart < bEnd && bStart < aEnd;
}

export function resolveRange(sourceMap, mdStart, mdEnd) {
  if (mdEnd <= mdStart) return [];

  for (const [ls, le] of sourceMap.locked || []) {
    if (rangesOverlap(mdStart, mdEnd, ls, le)) {
      throw new SourceMapError("locked", `edit [${mdStart},${mdEnd}) overlaps a locked range [${ls},${le})`, [mdStart, mdEnd]);
    }
  }

  const triples = [];
  let covered = 0;

  for (const block of sourceMap.blocks || []) {
    if (!rangesOverlap(mdStart, mdEnd, block.mdStart, block.mdEnd)) continue;
    for (const run of block.runs || []) {
      const s = Math.max(run.mdStart, mdStart);
      const e = Math.min(run.mdEnd, mdEnd);
      if (e <= s) continue;
      const dropStart = s - run.mdStart;
      triples.push({
        bodyPath: block.bodyPath,
        runIndex: run.runIndex,
        charStart: run.charOffset + dropStart,
        charEnd: run.charOffset + dropStart + (e - s),
      });
      covered += e - s;
    }
  }

  const requestedLen = mdEnd - mdStart;
  if (covered < requestedLen) {
    throw new SourceMapError(
      "synthetic",
      `edit [${mdStart},${mdEnd}) is not fully covered by document-text runs (resolved ${covered}/${requestedLen} chars)`,
      [mdStart, mdEnd]
    );
  }

  return triples;
}

// Edit-boundary snapping (spec §5.2): if an edit span's boundary falls inside a synthetic
// range (e.g. the LLM deleted bold text including the "**"), snap that boundary inward
// past the synthetic run. Locked ranges are a subset of synthetic (a locked placeholder is
// always also recorded as synthetic -- see ooxml/export.js's Composer.writeLocked), so a
// boundary that merely touches the edge of a locked island snaps fully past it too; an
// edit that still overlaps a locked range in its *interior* after snapping is untouched
// here and caught by resolveRange's unconditional locked check instead.
//
// Returns [start, end); start >= end means the requested span was entirely inside
// synthetic content ("entirely synthetic", spec §7 G4) -- callers must reject that rather
// than pass it to resolveRange (which treats an empty range as "nothing to resolve").
// resolveRange itself is unchanged and stays strict/non-snapping: this function trims the
// edges, resolveRange resolves (or throws on) what's left.
function snapStartForward(ranges, pos) {
  for (const [s, e] of ranges) {
    if (s > pos) break; // ranges are sorted; no later range can contain pos
    if (s <= pos && pos < e) return e;
  }
  return pos;
}
function snapEndBackward(ranges, pos) {
  for (const [s, e] of ranges) {
    if (s >= pos) break;
    if (s < pos && pos <= e) return s;
  }
  return pos;
}
export function snapBoundary(sourceMap, mdStart, mdEnd) {
  const synthetic = sourceMap.synthetic || [];
  let s = mdStart, e = mdEnd;
  for (;;) {
    const nextS = snapStartForward(synthetic, s);
    const nextE = snapEndBackward(synthetic, e);
    if (nextS === s && nextE === e) break;
    s = nextS;
    e = nextE;
  }
  if (e < s) e = s;
  return [s, e];
}

// Point-anchor resolution for insertions and bare point-comments (spec §4), which anchor a
// zero-width position rather than a span of pre-existing document text -- snapBoundary's
// span-trimming semantics aren't well-defined at an exact zero-width point (see M2 plan
// decision 1), so this is a separate, deliberately narrower resolver.
//
// M3b plan D2: the return shape is discriminated by `kind`:
//   - { kind: "run", bodyPath, runIndex, charOffset } -- an ordinary in-run point: pos
//     falls within (inclusive of both edges) some document-text run.
//   - { kind: "paragraphBoundary", bodyPath, edge: "before"|"after" } -- a whole-paragraph
//     point: pos sits at/within the gap between two blocks (or before the first/after the
//     last), matched against the exact block boundary rather than in-run geometry.
// Position alone can't disambiguate the two in every case (e.g. inserting at the very
// start of a no-prefix paragraph's first run == inserting a new paragraph immediately
// before it -- same md offset, structurally different DOM operations), so the caller
// passes opts.wholeParagraph (from criticmarkup/parse.js's D1 raw-token-shape flag) to
// pick the branch; this function does not infer it from pos.
//
// Throws SourceMapError("locked", ...) if pos sits strictly inside a locked range's
// interior (can't insert into the middle of a protected island) -- checked before either
// branch, since it applies regardless of wholeParagraph. Throws SourceMapError("synthetic",
// ...) if no run (ordinary case) or no aligned boundary (whole-paragraph case) covers pos.
export function resolvePoint(sourceMap, pos, opts = {}) {
  const wholeParagraph = !!opts.wholeParagraph;
  const blocks = sourceMap.blocks || [];

  for (const [ls, le] of sourceMap.locked || []) {
    if (ls < pos && pos < le) {
      throw new SourceMapError("locked", `position ${pos} is inside a locked range [${ls},${le})`, [pos, pos]);
    }
  }

  if (wholeParagraph) {
    if (!blocks.length) {
      throw new SourceMapError("synthetic", `position ${pos} does not align with a paragraph boundary (document has no blocks)`, [pos, pos]);
    }
    if (pos <= blocks[0].mdStart) {
      return { kind: "paragraphBoundary", bodyPath: blocks[0].bodyPath, edge: "before" };
    }
    if (pos >= blocks[blocks.length - 1].mdEnd) {
      return { kind: "paragraphBoundary", bodyPath: blocks[blocks.length - 1].bodyPath, edge: "after" };
    }
    for (let i = 0; i < blocks.length - 1; i++) {
      const a = blocks[i], b = blocks[i + 1];
      if (pos >= a.mdEnd && pos <= b.mdStart) {
        return { kind: "paragraphBoundary", bodyPath: a.bodyPath, edge: "after" };
      }
    }
    throw new SourceMapError(
      "synthetic",
      `position ${pos} does not align with any paragraph boundary for a whole-paragraph insert`,
      [pos, pos]
    );
  }

  for (const block of blocks) {
    if (pos < block.mdStart || pos > block.mdEnd) continue;
    for (const run of block.runs || []) {
      if (run.mdStart <= pos && pos <= run.mdEnd) {
        return { kind: "run", bodyPath: block.bodyPath, runIndex: run.runIndex, charOffset: run.charOffset + (pos - run.mdStart) };
      }
    }
  }
  throw new SourceMapError(
    "synthetic",
    `position ${pos} does not fall within any document-text run`,
    [pos, pos]
  );
}
