// Chunk mode (spec §6.4; two-phase split per architecture doc §7 / M4c step 2).
//
// The document under test here is built synthetically (not from a real .docx fixture):
// none of the committed fixtures (fixtures/generate.py) contain more than one top-level
// (H1) heading, so none of them actually exercise a multi-chunk split. Constructing the
// {markdown, sourceMap} pair directly -- the same way tests/validate.test.js's G5 suite
// builds its fakeSourceMap -- gets real multi-section coverage without touching the
// fixture corpus or its generation script.
import { describe, expect, it } from "vitest";
import { splitIntoChunks, translateEdits, validateChunked } from "../src/chunk.js";
import { validate } from "../src/validate.js";
import { mulberry32, buildValidResponse } from "./helpers/randomEdits.js";

const HEADER =
  "<!-- AutoReviewer export -->\n" +
  "<!-- CriticMarkup legend: {++ins++} {--del--} {~~old~>new~~} {==hl==}{>>comment<<} -->\n" +
  "<!-- no tracked changes or comments detected -->";

// Builds a {markdown, sourceMap} pair shaped like a real export.js output -- a leading
// 3-line header, then one or more "# Heading" + paragraph sections -- but with every
// mdStart/mdEnd/bodyPath/run hand-computed rather than derived from a .docx. Each
// section's heading is always level 1 ("# ") unless the caller overrides `level`.
function buildMultiSectionDoc(sections) {
  let markdown = HEADER;
  const blocks = [];
  let bodyIdx = 0;

  for (const section of sections) {
    markdown += "\n\n";
    const prefix = "#".repeat(section.level || 1) + " ";
    const headingStart = markdown.length;
    markdown += prefix + section.heading;
    blocks.push({
      mdStart: headingStart,
      mdEnd: markdown.length,
      kind: "heading",
      bodyPath: [bodyIdx++],
      runs: [{ mdStart: headingStart + prefix.length, mdEnd: markdown.length, runIndex: 0, charOffset: 0 }],
    });
    for (const para of section.paragraphs) {
      markdown += "\n\n";
      const pStart = markdown.length;
      markdown += para;
      blocks.push({
        mdStart: pStart,
        mdEnd: markdown.length,
        kind: "p",
        bodyPath: [bodyIdx++],
        runs: [{ mdStart: pStart, mdEnd: markdown.length, runIndex: 0, charOffset: 0 }],
      });
    }
  }

  return { markdown, sourceMap: { docHash: "sha256-test-multi-section", blocks, locked: [], synthetic: [] } };
}

// Deep-removes rawStart/rawEnd/rawStarts -- fields that index into a RESPONSE string, not
// the document -- before comparing a chunked result against an unchunked one. These are
// expected to differ by design: chunk mode never translates them (only mdPos/mdStart/mdEnd
// are document-coordinate fields), so a chunk-i>0 edit's rawStart is a position within that
// chunk's own pasted response text, not within the hypothetical whole concatenated response
// the unchunked baseline parses. Everything else -- type, translated md-offsets, anchor,
// text content, warnings, counts -- must match exactly.
function stripRaw(value) {
  if (Array.isArray(value)) return value.map(stripRaw);
  if (value && typeof value === "object") {
    const out = {};
    for (const [key, v] of Object.entries(value)) {
      if (key === "rawStart" || key === "rawEnd" || key === "rawStarts") continue;
      out[key] = stripRaw(v);
    }
    return out;
  }
  return value;
}

describe("chunk mode: splitting", () => {
  it("splits at each top-level heading; chunk 0 carries the document header", () => {
    const { markdown, sourceMap } = buildMultiSectionDoc([
      { heading: "Intro", paragraphs: ["First paragraph text here.", "Second paragraph text here."] },
      { heading: "Middle Section", paragraphs: ["Third paragraph of prose."] },
      { heading: "Final Section", paragraphs: ["Fourth paragraph of prose.", "Fifth paragraph of prose."] },
    ]);
    const chunks = splitIntoChunks(markdown, sourceMap);
    expect(chunks.length).toBe(3);
    expect(chunks[0].baseOffset).toBe(0);
    expect(chunks[0].exportedMarkdown.startsWith(HEADER)).toBe(true);
    for (let i = 1; i < chunks.length; i++) {
      expect(chunks[i].exportedMarkdown.startsWith("# ")).toBe(true);
      expect(chunks[i].sourceMap.blocks[0].mdStart).toBe(0);
    }
    // Chunks partition the document exactly -- no gap, no overlap, no reordering.
    expect(chunks.map((c) => c.exportedMarkdown).join("")).toBe(markdown);
  });

  it("does not split on H2-H6 -- only a literal single-'#' heading is a chunk boundary", () => {
    const { markdown, sourceMap } = buildMultiSectionDoc([
      { heading: "Top", paragraphs: ["Paragraph under the top heading."] },
      { heading: "Sub", level: 2, paragraphs: ["Paragraph under the level-2 sub-heading."] },
    ]);
    const chunks = splitIntoChunks(markdown, sourceMap);
    expect(chunks.length).toBe(1);
    expect(chunks[0].exportedMarkdown).toBe(markdown);
  });

  it("a document with no top-level heading at all yields exactly one chunk spanning the whole document", () => {
    const { markdown, sourceMap } = buildMultiSectionDoc([{ heading: "Only Section", paragraphs: ["Just one paragraph."] }]);
    sourceMap.blocks[0].kind = "p"; // demote the only heading so there is no H1 anywhere
    const chunks = splitIntoChunks(markdown, sourceMap);
    expect(chunks.length).toBe(1);
    expect(chunks[0].baseOffset).toBe(0);
    expect(chunks[0].exportedMarkdown).toBe(markdown);
  });

  it("chunk-ignorance: neither sourcemap.js nor ooxml/inject.js reference chunk or baseOffset", async () => {
    // Executable form of the M4c guardrail (architecture doc §7 item 5): chunk knowledge
    // must die at this file's own offset-translation boundary. Reading the two frozen
    // files' source directly (rather than shelling out to grep) keeps this check running
    // under Vitest like every other assertion here.
    const { readFileSync } = await import("node:fs");
    const path = await import("node:path");
    const { fileURLToPath } = await import("node:url");
    const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
    for (const rel of ["src/sourcemap.js", "src/ooxml/inject.js"]) {
      const text = readFileSync(path.join(root, rel), "utf8");
      expect(text, rel).not.toMatch(/chunk|baseOffset/i);
    }
  });
});

describe("chunk mode: offset translation", () => {
  it("adds baseOffset to every md-offset field (mdPos, mdStart, mdEnd) and to none of rawStart/rawEnd", () => {
    const edits = [
      { type: "ins", mdPos: 10, rawStart: 3, rawEnd: 8 },
      { type: "del", mdStart: 5, mdEnd: 9, rawStart: 1, rawEnd: 4 },
      { type: "comment", anchored: false, mdPos: 2, rawStart: 0, rawEnd: 6 },
    ];
    const translated = translateEdits(edits, 100);
    expect(translated[0]).toEqual({ type: "ins", mdPos: 110, rawStart: 3, rawEnd: 8 });
    expect(translated[1]).toEqual({ type: "del", mdStart: 105, mdEnd: 109, rawStart: 1, rawEnd: 4 });
    expect(translated[2]).toEqual({ type: "comment", anchored: false, mdPos: 102, rawStart: 0, rawEnd: 6 });
  });

  it("is a no-op for baseOffset 0 (chunk 0 never needs translation)", () => {
    const edits = [{ type: "ins", mdPos: 4, rawStart: 0, rawEnd: 5 }];
    expect(translateEdits(edits, 0)).toEqual(edits);
  });
});

describe("chunk mode: chunked-vs-unchunked equivalence oracle (architecture doc §7)", () => {
  const seed = Number(process.env.AR_FUZZ_SEED) || Date.now();
  const rng = mulberry32(seed);

  const { markdown, sourceMap } = buildMultiSectionDoc([
    {
      heading: "Introduction",
      paragraphs: [
        "This regulatory text discusses cost-benefit considerations for the proposed rule.",
        "Compliance timelines are addressed in the following sections of this document.",
      ],
    },
    {
      heading: "Economic Analysis",
      paragraphs: [
        "The analysis considers direct and indirect costs to affected carriers nationwide.",
        "Benefits accrue primarily to consumers through improved safety outcomes overall.",
      ],
    },
    {
      heading: "Conclusion",
      paragraphs: ["The agency recommends adoption of the rule as proposed without further changes."],
    },
  ]);
  const chunks = splitIntoChunks(markdown, sourceMap);

  it("sanity: this fixture actually splits into more than one chunk", () => {
    expect(chunks.length).toBe(3);
  });

  it(
    `~60 random valid per-chunk response sets: the two-phase chunked pipeline produces the same edits/warnings/counts ` +
      `(modulo chunk-local rawStart/rawEnd) as a single-pass validate() over the reassembled full response (seed=${seed})`,
    () => {
      const ITERATIONS = 60;
      for (let i = 0; i < ITERATIONS; i++) {
        const chunkResponses = chunks.map((c) => buildValidResponse(rng, c.exportedMarkdown, c.sourceMap));
        const fullResponse = chunkResponses.join("");

        const baseline = validate({ responseMarkdown: fullResponse, exportedMarkdown: markdown, sourceMap });
        const chunked = validateChunked({ chunks, chunkResponses, sourceMap });

        const ctx = `seed=${seed} iter=${i}`;
        expect(baseline.ok, `${ctx}: baseline unexpectedly failed -- gate=${baseline.gate} message=${baseline.message}`).toBe(
          true
        );
        expect(
          chunked.ok,
          `${ctx}: chunked unexpectedly failed -- gate=${chunked.gate} message=${chunked.message} chunkIndex=${chunked.chunkIndex}`
        ).toBe(true);

        expect(stripRaw(chunked.edits), ctx).toEqual(stripRaw(baseline.edits));
        expect(stripRaw(chunked.warnings), ctx).toEqual(stripRaw(baseline.warnings));
        expect(chunked.counts, ctx).toEqual(baseline.counts);
      }
    }
  );
});
