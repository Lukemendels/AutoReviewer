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
// decision 1), so this is a separate, deliberately narrower resolver:
//   - throws SourceMapError("locked", ...) if pos sits strictly inside a locked range's
//     interior (can't insert into the middle of a protected island);
//   - otherwise returns { bodyPath } for the block whose [mdStart, mdEnd] contains pos
//     (inclusive of both edges);
//   - throws SourceMapError("synthetic", ...) if pos falls in the gap between blocks (or
//     in the header/orphan-comments section, which have no blocks at all) -- this includes
//     the whole-paragraph-insert case (spec §4), whose injection semantics belong to M3
//     (paragraph-mark handling, spec §9.1 step 7) and aren't resolvable yet.
export function resolvePoint(sourceMap, pos) {
  for (const [ls, le] of sourceMap.locked || []) {
    if (ls < pos && pos < le) {
      throw new SourceMapError("locked", `position ${pos} is inside a locked range [${ls},${le})`, [pos, pos]);
    }
  }
  for (const block of sourceMap.blocks || []) {
    if (block.mdStart <= pos && pos <= block.mdEnd) {
      return { bodyPath: block.bodyPath };
    }
  }
  throw new SourceMapError(
    "synthetic",
    `position ${pos} does not fall within any block (whole-paragraph inserts are not resolvable until M3)`,
    [pos, pos]
  );
}
