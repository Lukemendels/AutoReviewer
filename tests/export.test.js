import { readFileSync, readdirSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { resolveRange, SourceMapError } from "../src/sourcemap.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

function loadFixture(name) {
  const buf = readFileSync(path.join(fixturesDir, name));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

const fixtureNames = readdirSync(fixturesDir).filter((f) => f.endsWith(".docx"));

async function exportFixture(name) {
  const bytes = loadFixture(name);
  return exportDocx(bytes, { DOMParserImpl: DOMParser, filename: name.replace(/\.docx$/, "") });
}

describe("fixture corpus", () => {
  it("found the expected fixtures", () => {
    expect(fixtureNames.sort()).toEqual(
      [
        "bold-italic.docx",
        "comments-threaded.docx",
        "fields-and-content-controls.docx",
        "headings-and-lists.docx",
        "hyperlinks-and-images.docx",
        "plain-paragraphs.docx",
        "stressor.docx",
        "tables.docx",
        "tracked-changes.docx",
      ].sort()
    );
  });
});

describe("export determinism", () => {
  for (const name of fixtureNames) {
    it(`${name}: same input produces identical markdown + source map`, async () => {
      const a = await exportFixture(name);
      const b = await exportFixture(name);
      expect(a.markdown).toBe(b.markdown);
      expect(a.sourceMap).toEqual(b.sourceMap);
      expect(a.counts).toEqual(b.counts);
    });
  }
});

describe("plain-paragraphs.docx (redline-free baseline)", () => {
  it("renders three plain paragraphs separated by blank lines, no CriticMarkup", async () => {
    const { markdown } = await exportFixture("plain-paragraphs.docx");
    const body = markdown.split("-->\n\n")[1];
    expect(body).toBe(
      "This is the first paragraph of a plain document.\n\n" +
        "This is the second paragraph, with nothing tracked.\n\n" +
        "A third paragraph exists so table/list neighbors in later fixtures have a plain-paragraph baseline to compare against.\n"
    );
    expect(body).not.toMatch(/\{[+~-]/);
  });

  it("header comment reports no tracked changes or comments", async () => {
    const { markdown } = await exportFixture("plain-paragraphs.docx");
    expect(markdown).toContain("<!-- no tracked changes or comments detected -->");
  });
});

describe("headings-and-lists.docx", () => {
  it("renders heading levels and bullet/numbered lists", async () => {
    const { markdown } = await exportFixture("headings-and-lists.docx");
    expect(markdown).toContain("# Top-Level Heading");
    expect(markdown).toContain("## A Subsection");
    expect(markdown).toContain("### A Sub-subsection");
    expect(markdown).toContain("- First bullet");
    expect(markdown).toContain("- First numbered item");
  });
});

describe("tables.docx", () => {
  it("renders a pipe-delimited table with header separator row", async () => {
    const { markdown } = await exportFixture("tables.docx");
    expect(markdown).toContain("| Column A | Column B | Column C |");
    expect(markdown).toContain("| --- | --- | --- |");
    expect(markdown).toContain("| Row 1, cell A | Row 1, cell B is a bit longer than the others | 1 |");
  });
});

describe("bold-italic.docx", () => {
  it("wraps bold/italic/bold+italic runs with correct markers", async () => {
    const { markdown } = await exportFixture("bold-italic.docx");
    expect(markdown).toContain("**bold text**");
    expect(markdown).toContain("*italic text*");
    expect(markdown).toContain("***bold italic text***");
  });
});

describe("hyperlinks-and-images.docx", () => {
  it("renders the hyperlink and a locked image placeholder", async () => {
    const { markdown } = await exportFixture("hyperlinks-and-images.docx");
    expect(markdown).toContain("[regulatory docket](https://example.gov/docket/12345)");
    expect(markdown).toMatch(/\[image: image\d*\.\w+\]/);
  });

  it("locks the image placeholder range", async () => {
    const { markdown, sourceMap } = await exportFixture("hyperlinks-and-images.docx");
    const idx = markdown.indexOf("[image:");
    expect(idx).toBeGreaterThan(-1);
    const end = markdown.indexOf("]", idx) + 1;
    expect(() => resolveRange(sourceMap, idx, end)).toThrow(SourceMapError);
    try {
      resolveRange(sourceMap, idx, end);
    } catch (err) {
      expect(err.kind).toBe("locked");
    }
  });
});

describe("tracked-changes.docx", () => {
  it("renders insertion, deletion, and substitution CriticMarkup", async () => {
    const { markdown, counts } = await exportFixture("tracked-changes.docx");
    expect(markdown).toContain("{++all covered carriers++}");
    expect(markdown).toContain("{--should be removed entirely--}");
    expect(markdown).toContain("{~~30 days~>60 days~~}");
    expect(counts.ins).toBe(1);
    expect(counts.del).toBe(1);
    expect(counts.sub).toBe(1);
  });

  it("treats pre-existing tracked-change text as synthetic, not document-text runs", async () => {
    const { markdown, sourceMap } = await exportFixture("tracked-changes.docx");
    const idx = markdown.indexOf("{++all covered carriers++}");
    expect(() => resolveRange(sourceMap, idx, idx + "{++all covered carriers++}".length)).toThrow(SourceMapError);
  });
});

describe("comments-threaded.docx", () => {
  it("renders a reply thread and a resolved tag", async () => {
    const { markdown } = await exportFixture("comments-threaded.docx");
    expect(markdown).toContain("Is this the right threshold?");
    expect(markdown).toContain("↳ Reviewer B");
    expect(markdown).toContain("Agreed -- flagged for legal review.");
    expect(markdown).toContain("[resolved]");
    expect(markdown).toContain("Please cite the authority here.");
  });
});

describe("fields-and-content-controls.docx", () => {
  it("renders a locked field placeholder using the field instruction name", async () => {
    const { markdown } = await exportFixture("fields-and-content-controls.docx");
    expect(markdown).toContain("⟦field: PAGE⟧");
  });

  it("renders the content control's inner text and locks it", async () => {
    const { markdown, sourceMap } = await exportFixture("fields-and-content-controls.docx");
    const idx = markdown.indexOf("TSA-2026-0042");
    expect(idx).toBeGreaterThan(-1);
    expect(() => resolveRange(sourceMap, idx, idx + "TSA-2026-0042".length)).toThrow(SourceMapError);
  });
});

describe("source map: blocks/runs resolve to themselves across the full corpus", () => {
  for (const name of fixtureNames) {
    it(`${name}: every registered run resolves back to its own bodyPath/runIndex/charOffset`, async () => {
      const { sourceMap } = await exportFixture(name);
      expect(sourceMap.blocks.length).toBeGreaterThan(0);
      let checked = 0;
      for (const block of sourceMap.blocks) {
        for (const run of block.runs) {
          const triples = resolveRange(sourceMap, run.mdStart, run.mdEnd);
          expect(triples).toEqual([
            {
              bodyPath: block.bodyPath,
              runIndex: run.runIndex,
              charStart: run.charOffset,
              charEnd: run.charOffset + (run.mdEnd - run.mdStart),
            },
          ]);
          checked++;
        }
      }
      expect(checked).toBeGreaterThan(0);
    });
  }

  it("plain-paragraphs.docx: resolveRange spans a whole block across multiple runs", async () => {
    const { sourceMap } = await exportFixture("plain-paragraphs.docx");
    const block = sourceMap.blocks[0];
    const triples = resolveRange(sourceMap, block.mdStart, block.mdEnd);
    expect(triples.length).toBeGreaterThan(0);
    const total = triples.reduce((sum, t) => sum + (t.charEnd - t.charStart), 0);
    expect(total).toBe(block.mdEnd - block.mdStart);
  });

  it("synthetic/locked ranges never overlap a block's own runs", async () => {
    for (const name of fixtureNames) {
      const { sourceMap } = await exportFixture(name);
      for (const block of sourceMap.blocks) {
        for (const run of block.runs) {
          for (const [ls, le] of sourceMap.locked) {
            const overlaps = run.mdStart < le && ls < run.mdEnd;
            expect(overlaps).toBe(false);
          }
        }
      }
    }
  });

  it("an offset inside the synthetic header comment cannot be resolved", async () => {
    const { sourceMap } = await exportFixture("plain-paragraphs.docx");
    expect(() => resolveRange(sourceMap, 0, 10)).toThrow(SourceMapError);
    try {
      resolveRange(sourceMap, 0, 10);
    } catch (err) {
      expect(err.kind).toBe("synthetic");
    }
  });

  it("docHash is a sha256- prefixed hex digest of the original bytes", async () => {
    const { sourceMap } = await exportFixture("plain-paragraphs.docx");
    expect(sourceMap.docHash).toMatch(/^sha256-[0-9a-f]{64}$/);
  });
});

describe("stressor.docx (perf/determinism at scale)", () => {
  it("exports without error and produces a large, internally consistent source map", async () => {
    const { markdown, sourceMap, counts } = await exportFixture("stressor.docx");
    expect(markdown.length).toBeGreaterThan(10000);
    expect(sourceMap.blocks.length).toBeGreaterThan(100);
    expect(counts.ins).toBeGreaterThan(0);
  });
});
