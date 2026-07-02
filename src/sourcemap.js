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
