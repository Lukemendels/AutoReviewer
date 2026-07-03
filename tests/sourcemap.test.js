import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { resolvePoint, resolveRange, snapBoundary, SourceMapError } from "../src/sourcemap.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

function loadDocx(name) {
  const buf = readFileSync(path.join(fixturesDir, `${name}.docx`));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

// "**bold**": synthetic "**" at [0,2) and [6,8), real run text "bold" at [2,6).
const boldSourceMap = {
  blocks: [{ mdStart: 0, mdEnd: 8, kind: "p", bodyPath: [0], runs: [{ mdStart: 2, mdEnd: 6, runIndex: 0, charOffset: 0 }] }],
  synthetic: [[0, 2], [6, 8]],
  locked: [],
};

describe("snapBoundary: pure boundary math", () => {
  it("snaps a full-token span inward past leading and trailing emphasis markers", () => {
    expect(snapBoundary(boldSourceMap, 0, 8)).toEqual([2, 6]);
  });

  it("leaves an already run-aligned span untouched", () => {
    expect(snapBoundary(boldSourceMap, 2, 6)).toEqual([2, 6]);
  });

  it("snaps only the boundary that actually lands inside synthetic content", () => {
    expect(snapBoundary(boldSourceMap, 1, 6)).toEqual([2, 6]); // start inside "**"
    expect(snapBoundary(boldSourceMap, 2, 7)).toEqual([2, 6]); // end inside "**"
  });

  it("collapses a span entirely inside synthetic content to empty (start === end)", () => {
    const [s, e] = snapBoundary(boldSourceMap, 0, 2);
    expect(s).toBe(e);
  });

  it("is a no-op on a span with no synthetic overlap at all", () => {
    expect(snapBoundary(boldSourceMap, 3, 5)).toEqual([3, 5]);
  });
});

describe("snapBoundary + resolveRange integration", () => {
  it("a snapped span resolves cleanly to the underlying run", () => {
    const [s, e] = snapBoundary(boldSourceMap, 0, 8);
    expect(resolveRange(boldSourceMap, s, e)).toEqual([{ bodyPath: [0], runIndex: 0, charStart: 0, charEnd: 4 }]);
  });

  it("a fully-synthetic span, once snapped to empty, must not be handed to resolveRange as a real edit", () => {
    const [s, e] = snapBoundary(boldSourceMap, 0, 2);
    expect(s).toBe(e);
    // resolveRange treats an empty range as "nothing to resolve", not an error -- callers
    // (M2's G4) are responsible for rejecting s === e as "entirely synthetic" themselves.
    expect(resolveRange(boldSourceMap, s, e)).toEqual([]);
  });

  it("locked content in a span's interior (not at its edges) is unaffected by snapping and still rejected by resolveRange", async () => {
    const { markdown, sourceMap } = await exportDocx(loadDocx("hyperlinks-and-images"), {
      DOMParserImpl: DOMParser,
      filename: "hyperlinks-and-images",
    });
    const before = markdown.indexOf("See the ");
    const imgStart = markdown.indexOf("[image:");
    const after = markdown.indexOf("Text after the image.");
    const afterEnd = after + "Text after the image.".length;
    expect(before).toBeGreaterThanOrEqual(0);
    expect(imgStart).toBeGreaterThan(before);
    expect(afterEnd).toBeGreaterThan(imgStart);

    // Both boundaries sit inside real document text (not synthetic), so snapping is a no-op...
    const [s, e] = snapBoundary(sourceMap, before, afterEnd);
    expect([s, e]).toEqual([before, afterEnd]);
    // ...yet the span's interior still contains the whole locked image placeholder, so
    // resolveRange must still reject it.
    expect(() => resolveRange(sourceMap, s, e)).toThrow(SourceMapError);
    try {
      resolveRange(sourceMap, s, e);
    } catch (err) {
      expect(err.kind).toBe("locked");
    }
  });

  it("a boundary that merely touches the edge of a locked field placeholder snaps fully past it", async () => {
    const { markdown, sourceMap } = await exportDocx(loadDocx("fields-and-content-controls"), {
      DOMParserImpl: DOMParser,
      filename: "fields-and-content-controls",
    });
    const fieldStart = markdown.indexOf("⟦field:");
    const fieldEnd = markdown.indexOf("⟧", fieldStart) + 1;
    // Request a span starting one char into the placeholder through one char past it --
    // the start boundary is inside the locked/synthetic run, so it must snap forward past
    // the whole placeholder rather than partially resolving into it.
    const [s] = snapBoundary(sourceMap, fieldStart + 1, fieldEnd + 1);
    expect(s).toBe(fieldEnd);
  });
});

// D2 (M3b plan): resolvePoint's discriminated result. boldSourceMap's single block spans
// [0,8) with a document-text run at [2,6) ("bold"), synthetic "**" at the edges.
describe("resolvePoint: discriminated run/paragraphBoundary result (D2)", () => {
  it("an ordinary in-run point resolves to {kind:'run', runIndex, charOffset}", () => {
    expect(resolvePoint(boldSourceMap, 4)).toEqual({ kind: "run", bodyPath: [0], runIndex: 0, charOffset: 2 });
  });

  it("a point at a run's own edge (still inside the block, not whole-paragraph) resolves to charOffset 0 or the run's full length", () => {
    expect(resolvePoint(boldSourceMap, 2)).toEqual({ kind: "run", bodyPath: [0], runIndex: 0, charOffset: 0 });
    expect(resolvePoint(boldSourceMap, 6)).toEqual({ kind: "run", bodyPath: [0], runIndex: 0, charOffset: 4 });
  });

  it("a non-whole-paragraph point that lands in synthetic scaffolding (not inside any run) is rejected as G3 (synthetic)", () => {
    expect(() => resolvePoint(boldSourceMap, 0)).toThrow(SourceMapError);
    try {
      resolvePoint(boldSourceMap, 0);
    } catch (err) {
      expect(err.kind).toBe("synthetic");
    }
  });

  it("a whole-paragraph point exactly at a no-prefix block's own start resolves to {kind:'paragraphBoundary'}, even though the exact same position also resolves as an ordinary in-run point", () => {
    // A no-prefix paragraph: the block's own mdStart coincides exactly with its first
    // run's mdStart (no synthetic heading/list prefix in between) -- the plan's own "why
    // not position-based" example. Position 0 is simultaneously a valid in-run point
    // (charOffset 0) and a valid paragraph boundary; opts.wholeParagraph is what
    // disambiguates (see the D2 comment in sourcemap.js), not geometry alone.
    const noPrefixMap = {
      blocks: [{ mdStart: 0, mdEnd: 5, kind: "p", bodyPath: [0], runs: [{ mdStart: 0, mdEnd: 5, runIndex: 0, charOffset: 0 }] }],
      synthetic: [],
      locked: [],
    };
    expect(resolvePoint(noPrefixMap, 0)).toEqual({ kind: "run", bodyPath: [0], runIndex: 0, charOffset: 0 });
    expect(resolvePoint(noPrefixMap, 0, { wholeParagraph: true })).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "before" });
  });

  it("a whole-paragraph point between two adjacent blocks resolves to {edge:'after'} on the earlier block", () => {
    const twoBlockMap = {
      blocks: [
        { mdStart: 0, mdEnd: 5, kind: "p", bodyPath: [0], runs: [{ mdStart: 0, mdEnd: 5, runIndex: 0, charOffset: 0 }] },
        { mdStart: 7, mdEnd: 12, kind: "p", bodyPath: [1], runs: [{ mdStart: 7, mdEnd: 12, runIndex: 0, charOffset: 0 }] },
      ],
      synthetic: [],
      locked: [],
    };
    // Gap is [5,7) (e.g. "\n\n"); a whole-paragraph insertion token spliced into that gap
    // resolves to a point strictly inside it once the token itself is excised (D1).
    expect(resolvePoint(twoBlockMap, 6, { wholeParagraph: true })).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "after" });
    expect(resolvePoint(twoBlockMap, 5, { wholeParagraph: true })).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "after" });
    expect(resolvePoint(twoBlockMap, 7, { wholeParagraph: true })).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "after" });
  });

  it("a whole-paragraph point before the first block or after the last resolves with the corresponding edge", () => {
    expect(resolvePoint(boldSourceMap, 0, { wholeParagraph: true })).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "before" });
    expect(resolvePoint(boldSourceMap, 8, { wholeParagraph: true })).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "after" });
  });

  it("still rejects a position strictly inside a locked range's interior, whole-paragraph or not", async () => {
    const { markdown, sourceMap } = await exportDocx(loadDocx("fields-and-content-controls"), {
      DOMParserImpl: DOMParser,
      filename: "fields-and-content-controls",
    });
    const fieldStart = markdown.indexOf("⟦field:");
    const fieldEnd = markdown.indexOf("⟧", fieldStart) + 1;
    const insidePos = fieldStart + 2;
    expect(() => resolvePoint(sourceMap, insidePos)).toThrow(SourceMapError);
    expect(() => resolvePoint(sourceMap, insidePos, { wholeParagraph: true })).toThrow(SourceMapError);
    try {
      resolvePoint(sourceMap, insidePos);
    } catch (err) {
      expect(err.kind).toBe("locked");
    }
  });
});
