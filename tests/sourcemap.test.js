import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { resolveRange, snapBoundary, SourceMapError } from "../src/sourcemap.js";

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
