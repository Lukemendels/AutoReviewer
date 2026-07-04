// Unit-level tests of inject.js's run-location/run-splitting primitives directly (spec
// §9.1 steps 1-2; M3b plan's "run-location + run-splitting algorithm"), independent of
// the full injectEdits pipeline. Written and red before ooxml/inject.js has a real
// implementation.
import { describe, expect, it } from "vitest";
import { DOMParser } from "../tests/helpers/docx.js";
import { loadDocumentXmlDom, exportFixture } from "./helpers/docx.js";
import { resolveRange } from "../src/sourcemap.js";
import { NS } from "../src/ooxml/parse.js";
import { locateParagraph, locateRun, splitRun } from "../src/ooxml/inject.js";

function parseFragment(xml) {
  const wrapped = `<root xmlns:w="${NS.w}">${xml}</root>`;
  const doc = new DOMParser().parseFromString(wrapped, "application/xml");
  return { doc, root: doc.documentElement };
}

function serialize(el) {
  // Minimal, dependency-free serializer for assertions (avoids pulling in
  // ooxml/serialize.js, which is a different milestone slice) -- just enough of XML to
  // make node-shape assertions legible in test failure output.
  if (el.nodeType === 3) return el.textContent;
  let s = `<${el.tagName}`;
  for (const attr of [...(el.attributes || [])]) s += ` ${attr.name}="${attr.value}"`;
  s += ">";
  for (const c of [...el.childNodes]) s += serialize(c);
  s += `</${el.tagName}>`;
  return s;
}

function runText(el) {
  const t = [...el.children].find((c) => c.localName === "t");
  return t ? t.textContent : null;
}
function tHasPreserve(el) {
  const t = [...el.children].find((c) => c.localName === "t");
  return t ? t.getAttribute("xml:space") === "preserve" : false;
}

describe("locateParagraph: replicates walkBodyTracked's bodyIdx counting", () => {
  it("finds the first top-level paragraph by bodyPath [0]", async () => {
    const { body } = await loadDocumentXmlDom("plain-paragraphs");
    const p = locateParagraph(body, [0]);
    expect(p.localName).toBe("p");
    expect(p.getElementsByTagName("w:t")[0].textContent).toBe("This is the first paragraph of a plain document.");
  });

  it("finds the second and third top-level paragraphs by index, counting every body child", async () => {
    const { body } = await loadDocumentXmlDom("plain-paragraphs");
    const p1 = locateParagraph(body, [1]);
    const p2 = locateParagraph(body, [2]);
    expect(p1.getElementsByTagName("w:t")[0].textContent).toBe("This is the second paragraph, with nothing tracked.");
    expect(p2.getElementsByTagName("w:t")[0].textContent).toMatch(/^A third paragraph exists/);
  });

  it("resolves a table-cell bodyPath [bodyIdx, rowIdx, cellIdx, pIdx]", async () => {
    const { body } = await loadDocumentXmlDom("tables");
    // tables.docx: paragraph "A table follows." at bodyPath [0], then the table.
    const tableBodyIdx = 1;
    const p = locateParagraph(body, [tableBodyIdx, 0, 0, 0]);
    expect(p.localName).toBe("p");
    expect(p.getElementsByTagName("w:t")[0].textContent).toBe("Column A");
  });

  it("throws a clear error for an out-of-range bodyPath", async () => {
    const { body } = await loadDocumentXmlDom("plain-paragraphs");
    expect(() => locateParagraph(body, [999])).toThrow();
  });
});

describe("locateRun: replicates buildSegments' exact counting order", () => {
  it("returns the direct-child w:r for each runIndex the source map recorded, in order", async () => {
    const { sourceMap } = await exportFixture("bold-italic");
    const { body } = await loadDocumentXmlDom("bold-italic");
    const p = locateParagraph(body, sourceMap.blocks[0].bodyPath);
    const expectedTexts = [
      "Plain text, then ",
      "bold text",
      ", then ",
      "italic text",
      ", then ",
      "bold italic text",
      ", then plain again.",
    ];
    for (let i = 0; i < expectedTexts.length; i++) {
      const run = locateRun(p, i);
      expect(run, `runIndex ${i}`).toBeTruthy();
      expect(run.localName).toBe("r");
      expect(runText(run)).toBe(expectedTexts[i]);
    }
  });

  it("returns null for a runIndex beyond the paragraph's run count", async () => {
    const { body } = await loadDocumentXmlDom("plain-paragraphs");
    const p = locateParagraph(body, [0]);
    expect(locateRun(p, 999)).toBeNull();
  });

  it("skips runs nested inside w:hyperlink for direct-child indexing purposes (never a target) but still advances the counter", async () => {
    const { sourceMap } = await exportFixture("hyperlinks-and-images");
    const { body } = await loadDocumentXmlDom("hyperlinks-and-images");
    // First paragraph: run 0 "See the ", hyperlink run "regulatory docket" (counted, never
    // a direct-child target), run 2 " for background." -- so runIndex 1 must resolve to
    // nothing as a *direct* child, while runIndex 2 still lands on the right run despite
    // the hyperlink's nested run having consumed index 1.
    const p = locateParagraph(body, sourceMap.blocks[0].bodyPath);
    const run0 = locateRun(p, 0);
    const run2 = locateRun(p, 2);
    expect(runText(run0)).toBe("See the ");
    expect(runText(run2)).toBe(" for background.");
  });
});

describe("splitRun: peel-from-the-right primitive", () => {
  it("full-range split (charStart=0, charEnd=len) produces only a core piece, run removed from its parent", () => {
    const { doc, root } = parseFragment('<w:p><w:r><w:t>hello</w:t></w:r></w:p>');
    const p = root.firstChild;
    const run = p.firstChild;
    const result = splitRun(doc, run, 0, 5);
    expect(result.before).toBeNull();
    expect(result.after).toBeNull();
    expect(result.core).toBeTruthy();
    expect(runText(result.core)).toBe("hello");
    expect([...p.children]).toEqual([result.core]);
    expect(run.parentNode).toBeNull();
  });

  it("mid-run split produces before/core/after in document order, each a distinct w:r", () => {
    const { doc, root } = parseFragment('<w:p><w:r><w:t>hello world</w:t></w:r></w:p>');
    const p = root.firstChild;
    const run = p.firstChild;
    const result = splitRun(doc, run, 2, 7);
    expect(runText(result.before)).toBe("he");
    expect(runText(result.core)).toBe("llo w");
    expect(runText(result.after)).toBe("orld");
    expect([...p.children]).toEqual([result.before, result.core, result.after]);
  });

  it("zero-width split (charStart === charEnd) produces before/after with no core -- the insertion-point case", () => {
    const { doc, root } = parseFragment('<w:p><w:r><w:t>hello</w:t></w:r></w:p>');
    const p = root.firstChild;
    const run = p.firstChild;
    const result = splitRun(doc, run, 3, 3);
    expect(result.core).toBeNull();
    expect(runText(result.before)).toBe("hel");
    expect(runText(result.after)).toBe("lo");
    expect([...p.children]).toEqual([result.before, result.after]);
  });

  it("leaves untouched siblings after the split run exactly where they were (uses a captured reference node, not appendChild)", () => {
    const { doc, root } = parseFragment('<w:p><w:r><w:t>first</w:t></w:r><w:r><w:t>second</w:t></w:r></w:p>');
    const p = root.firstChild;
    const firstRun = p.firstChild;
    const secondRunOriginal = p.lastChild;
    const result = splitRun(doc, firstRun, 0, 5);
    expect([...p.children]).toEqual([result.core, secondRunOriginal]);
    expect(runText(secondRunOriginal)).toBe("second");
  });

  it("deep-copies rPr onto every split piece", () => {
    const { doc, root } = parseFragment('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>hello world</w:t></w:r></w:p>');
    const p = root.firstChild;
    const run = p.firstChild;
    const result = splitRun(doc, run, 2, 7);
    for (const piece of [result.before, result.core, result.after]) {
      const rPr = [...piece.children].find((c) => c.localName === "rPr");
      expect(rPr, serialize(piece)).toBeTruthy();
      expect([...rPr.children].some((c) => c.localName === "b")).toBe(true);
    }
    // rPr clones are independent nodes, not shared references, so mutating one later
    // (e.g. injectEdits' del/ins formatting changes) can never leak across pieces.
    const beforeRPr = [...result.before.children].find((c) => c.localName === "rPr");
    const coreRPr = [...result.core.children].find((c) => c.localName === "rPr");
    expect(beforeRPr).not.toBe(coreRPr);
  });

  it('sets xml:space="preserve" on a piece whose own text starts or ends with whitespace, and omits it otherwise', () => {
    const { doc, root } = parseFragment('<w:p><w:r><w:t>Plain text, then </w:t></w:r></w:p>');
    const p = root.firstChild;
    const run = p.firstChild;
    // Split "Plain text, then " into "Plain" | " text, then " -- the second piece starts
    // with a space and must be marked preserve; the first has no leading/trailing ws.
    const result = splitRun(doc, run, 0, 5);
    expect(runText(result.core)).toBe("Plain");
    expect(runText(result.after)).toBe(" text, then ");
    expect(tHasPreserve(result.core)).toBe(false);
    expect(tHasPreserve(result.after)).toBe(true);
  });

  it("multi-run del span (resolveRange's per-run triples): each run is split independently at its own local offsets", () => {
    // Two plain (unformatted) runs, contiguous in md-space with no synthetic gap between
    // them ("Hello " + "world" == "Hello world"), built directly (like
    // sourcemap.test.js's boldSourceMap) rather than from a fixture -- every real fixture
    // with adjacent multi-run paragraphs happens to separate its runs with a synthetic
    // emphasis marker (see bold-italic.docx), which makes the *interior* of a span
    // spanning them uncoverable by resolveRange (correctly -- G3's job, not this test's).
    const { doc, root } = parseFragment('<w:p><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>');
    const p = root.firstChild;
    const sourceMap = {
      blocks: [
        {
          mdStart: 0,
          mdEnd: 11,
          kind: "p",
          bodyPath: [0],
          runs: [
            { mdStart: 0, mdEnd: 6, runIndex: 0, charOffset: 0 },
            { mdStart: 6, mdEnd: 11, runIndex: 1, charOffset: 0 },
          ],
        },
      ],
      synthetic: [],
      locked: [],
    };

    const triples = resolveRange(sourceMap, 2, 9); // "llo " (run0) + "wor" (run1)
    expect(triples).toEqual([
      { bodyPath: [0], runIndex: 0, charStart: 2, charEnd: 6 },
      { bodyPath: [0], runIndex: 1, charStart: 0, charEnd: 3 },
    ]);

    // Process right-to-left (higher runIndex first), matching the algorithm's ordering.
    const ordered = [...triples].sort((a, b) => b.runIndex - a.runIndex);
    const results = {};
    for (const triple of ordered) {
      const run = locateRun(p, triple.runIndex);
      results[triple.runIndex] = splitRun(doc, run, triple.charStart, triple.charEnd);
    }
    expect(runText(results[0].before)).toBe("He");
    expect(runText(results[0].core)).toBe("llo ");
    expect(results[0].after).toBeNull();
    expect(results[1].before).toBeNull();
    expect(runText(results[1].core)).toBe("wor");
    expect(runText(results[1].after)).toBe("ld");

    expect([...p.children]).toEqual([results[0].before, results[0].core, results[1].core, results[1].after]);
  });
});
