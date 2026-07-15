// The gate fuzz suite (spec §13 item 4): every hand-crafted corruption below must be
// caught by its correctly-named gate, and nothing else. Valid/corrupted responses are
// built by splicing CriticMarkup tokens into the REAL output of exportDocx() on clean
// fixtures (no pre-existing tracked changes/comments baked into the export -- a response
// echoing pre-existing CriticMarkup verbatim raises its own G2 question, tracked
// separately and out of scope here), so the header/structure the response must echo
// verbatim is never hand-typed and can't silently drift from what export.js really emits.
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { validate } from "../src/validate.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

function loadDocx(name) {
  const buf = readFileSync(path.join(fixturesDir, `${name}.docx`));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}
async function exportFixture(name) {
  return exportDocx(loadDocx(name), { DOMParserImpl: DOMParser, filename: name });
}
// Splices a unique substring `target` to `replacement` in `markdown`, failing loudly if
// `target` isn't found exactly once (keeps every fuzz case's intent legible and safe from
// silent no-ops if a fixture's exported text ever changes).
function withEdit(markdown, target, replacement) {
  const first = markdown.indexOf(target);
  const last = markdown.lastIndexOf(target);
  if (first === -1) throw new Error(`withEdit: target not found: ${JSON.stringify(target)}`);
  if (first !== last) throw new Error(`withEdit: target not unique: ${JSON.stringify(target)}`);
  return markdown.slice(0, first) + replacement + markdown.slice(first + target.length);
}

describe("a fully valid response passes every gate", () => {
  it("plain-paragraphs.docx with one sub, one del, one ins, one bare comment", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    let response = exported;
    response = withEdit(response, "first", "{~~first~>initial~~}");
    response = withEdit(response, "second", "{--second--}");
    response = withEdit(response, "third paragraph", "{++really ++}third paragraph");
    response = withEdit(response, "compare against.", "compare against.{>>Double-check this framing.<<}");

    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(true);
    expect(result.edits.map((e) => e.type)).toEqual(["sub", "del", "ins", "comment"]);
    for (const edit of result.edits) expect(edit.anchor).toBeTruthy();
    expect(result.warnings).toEqual([]);
  });
});

describe("G1 -- grammar", () => {
  it("rejects nesting", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(exported, "first", "{~~fir{++st++}~>initial~~}");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G1");
  });

  it("rejects a token whose span crosses a block boundary", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(
      exported,
      ".\n\nThis is the second",
      "{--.\n\nThis is the second--}"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G1");
    expect(result.message).toMatch(/block/i);
  });

  it("rejects a highlight not immediately followed by a comment", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(exported, "first paragraph", "{==first paragraph==} {>>note<<}");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G1");
  });

  // Hotfix regression: a newline embedded in an insertion's own content passes tokenize/
  // strip/anchor-resolution untouched (none of them care what characters a token's content
  // holds), but injectEdits puts the whole string into one w:t verbatim -- the newline
  // lands as a literal character Word renders as a stray space, not a line break or a new
  // paragraph. Only the D1 whole-paragraph shape (token alone on its own line) is a
  // sanctioned way to add a new paragraph.
  it("rejects a newline embedded inside a mid-paragraph insertion's own text", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(exported, "third paragraph", "{++really\nreally ++}third paragraph");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G1");
    expect(result.message).toMatch(/newline/i);
    expect(result.message, "repair message must teach the D1 shape by example").toMatch(/alone on its own line/i);
    expect(result.message).toContain("{++New paragraph text.++}");
  });

  it("rejects a newline embedded inside a substitution's new text", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(exported, "first", "{~~first~>ini\ntial~~}");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G1");
    expect(result.message).toMatch(/newline/i);
  });

  it("does NOT reject a real D1 whole-paragraph insert, whose own token content has no embedded newline", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(
      exported,
      "document.\n\nThis is the second",
      "document.\n{++A whole new paragraph.++}\nThis is the second"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
  });
});

describe("G2 -- fidelity (paraphrase drift)", () => {
  it("rejects text changed outside any CriticMarkup token", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(exported, "This is the second paragraph", "This is now the second paragraph");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G2");
    // Cheap first-divergence diagnostics, not an eager full diff (perf: see validate.js's
    // G2 comment and tests/validate.perf.test.js) -- the UI computes the full word-level
    // diff lazily, on demand, from diffInputs.
    expect(result.firstDivergence).toBeTruthy();
    expect(result.firstDivergence.offset).toBeGreaterThanOrEqual(0);
    expect(result.diffInputs).toEqual({ a: expect.any(String), b: exported });
    expect(typeof result.repairPrompt).toBe("string");
    expect(result.repairPrompt.length).toBeGreaterThan(0);
  });

  it("catches drift in the header itself (the export header participates in G2 -- see issue #10)", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = exported.replace("no tracked changes or comments detected", "no changes detected");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G2");
  });

  // D7 (M3b plan): the header exemption is content-anchored and fails closed -- a response
  // that doesn't open with the export's own verbatim header is rejected as a G2-class
  // fabrication failure, the same way any other undetected drift is.
  it("D7: fails closed when the response is missing the header entirely", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const bodyOnly = exported.slice(exported.indexOf("This is the first paragraph"));
    const result = validate({ responseMarkdown: bodyOnly, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G2");
    expect(result.message).toMatch(/header/i);
  });

  it("D7: fails closed when the header is present but not at the very start of the response", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = "Some preamble the model added.\n" + exported;
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G2");
    expect(result.message).toMatch(/header/i);
  });

  // M4d PR-4 (F-6): a trailing-newline-only divergence gets an extra, specific sentence --
  // the generic G2 message alone is easy to misread when the first-divergence context
  // window shows nothing but whitespace.
  describe("trailing-newline-only divergence gets a clarifying message", () => {
    it("response missing the final trailing newline", async () => {
      const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
      expect(exported.endsWith("\n")).toBe(true);
      const response = exported.slice(0, -1);
      const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
      expect(result.ok).toBe(false);
      expect(result.gate).toBe("G2");
      expect(result.message).toMatch(/missing \(or adds\) a newline at the very end/);
      expect(result.message).toMatch(/last character before the closing fence must match the source exactly/);
    });

    it("response with an extra trailing newline", async () => {
      const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
      const response = exported + "\n";
      const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
      expect(result.ok).toBe(false);
      expect(result.gate).toBe("G2");
      expect(result.message).toMatch(/missing \(or adds\) a newline at the very end/);
    });

    it("does NOT append the clarifying sentence for an ordinary mid-document paraphrase drift", async () => {
      const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
      const response = withEdit(exported, "This is the second paragraph", "This is now the second paragraph");
      const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
      expect(result.ok).toBe(false);
      expect(result.gate).toBe("G2");
      expect(result.message).not.toMatch(/missing \(or adds\) a newline/);
    });
  });
});

describe("G3 -- anchor resolution", () => {
  it("rejects a deletion spanning into a non-locked synthetic interior (a hyperlink mid-span)", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("hyperlinks-and-images");
    const response = withEdit(
      exported,
      "See the [regulatory docket](https://example.gov/docket/12345) for",
      "{--See the [regulatory docket](https://example.gov/docket/12345) for--}"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G3");
  });

});

describe("whole-paragraph insertion (M3b: D1+D2 make this resolvable via resolvePoint's paragraphBoundary kind)", () => {
  it("a whole-paragraph insertion between two existing blocks now resolves successfully", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(
      exported,
      "document.\n\nThis is the second",
      "document.\n{++A whole new paragraph.++}\nThis is the second"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(true);
    const ins = result.edits.find((e) => e.type === "ins");
    expect(ins.wholeParagraph).toBe(true);
    expect(ins.anchor).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "after" });
  });

  it("an ordinary (non-whole-paragraph) insertion mid-sentence still resolves to an in-run point, not a paragraph boundary", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    const response = withEdit(exported, "third paragraph", "{++really ++}third paragraph");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(true);
    const ins = result.edits.find((e) => e.type === "ins");
    expect(ins.wholeParagraph).toBe(false);
    expect(ins.anchor.kind).toBe("run");
  });
});

describe("G4 -- protection", () => {
  it("rejects a deletion whose entire span is a locked placeholder (collapses to empty on snap)", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("hyperlinks-and-images");
    const response = withEdit(exported, "[image: image1.png]", "{--[image: image1.png]--}");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G4");
  });

  it("rejects a deletion whose span reaches partway into a locked field placeholder's interior", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("fields-and-content-controls");
    const response = withEdit(
      exported,
      "reference: ⟦field: PAGE⟧ of the",
      "{--reference: ⟦field: PAGE⟧ of the--}"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(false);
    expect(result.gate).toBe("G4");
  });
});

describe("G5 -- sanity report (warns, never blocks)", () => {
  // Built directly (not from a real export) to control word counts precisely without
  // fighting G1's block-crossing constraint.
  const longText = Array.from({ length: 60 }, (_, i) => `word${i}`).join(" ");
  const fakeExported = `Intro. ${longText} Outro.`;
  const fakeSourceMap = {
    docHash: "sha256-test",
    blocks: [
      {
        mdStart: 0,
        mdEnd: fakeExported.length,
        kind: "p",
        bodyPath: [0],
        runs: [{ mdStart: 0, mdEnd: fakeExported.length, runIndex: 0, charOffset: 0 }],
      },
    ],
    synthetic: [],
    locked: [],
  };

  it("flags an oversized deletion but still passes", () => {
    const response = withEdit(fakeExported, longText, `{--${longText}--}`);
    const result = validate({ responseMarkdown: response, exportedMarkdown: fakeExported, sourceMap: fakeSourceMap });
    expect(result.ok).toBe(true);
    expect(result.warnings.some((w) => w.code === "oversized-deletion")).toBe(true);
  });

  it("flags duplicate comment text but still passes", async () => {
    const { markdown: exported, sourceMap } = await exportFixture("plain-paragraphs");
    let response = exported;
    response = withEdit(response, "first paragraph of a plain document.", "first paragraph of a plain document.{>>Please verify.<<}");
    response = withEdit(response, "nothing tracked.", "nothing tracked.{>>Please verify.<<}");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported, sourceMap });
    expect(result.ok).toBe(true);
    expect(result.warnings.some((w) => w.code === "duplicate-comment")).toBe(true);
  });
});

describe("sentinel validation (M6a)", () => {
  it("validates a byte-perfect echo of a sentinelized document with zero edits", async () => {
    const bytes = loadDocx("comments-threaded-nested");
    const exported = await exportDocx(bytes, { DOMParserImpl: DOMParser, filename: "comments-threaded-nested", sentinel: true });
    const result = validate({ responseMarkdown: exported.markdown, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok).toBe(true);
    expect(result.edits).toHaveLength(0);
  });
});
