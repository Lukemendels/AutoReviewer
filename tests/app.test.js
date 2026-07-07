// src/ui/app.js's Inject wiring (M3b plan's wiring section). Exercises buildReviewedDocx
// -- the pure, DOM-download-free core of the flow -- against a real fixture end to end.
// Runs in the default (node) environment, passing @xmldom/xmldom explicitly: happy-dom's
// DOMParser does not reliably parse a real, namespace-heavy document.xml (verified
// directly -- it silently drops the entire document body, parsing only the root element's
// declarations), so it isn't a safe stand-in for XML parsing fidelity the way it is for
// ratify.js's plain-HTML-DOM tests. The browser build itself calls buildReviewedDocx with
// no Impl arguments, falling back to the real native globals, which don't have this gap.
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { validate } from "../src/validate.js";
import { createRatificationState } from "../src/ui/ratify.js";
import { buildReviewedDocx } from "../src/ui/app.js";
import { unzip, readEntry } from "../src/zip/reader.js";
import { buildAuditRecord } from "../src/audit.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

function loadDocxBytes(name) {
  const buf = readFileSync(path.join(fixturesDir, `${name}.docx`));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

describe("buildReviewedDocx: the Inject flow's pure core", () => {
  it("produces a re-openable docx with the accepted edits injected", async () => {
    const docxBytes = loadDocxBytes("plain-paragraphs");
    const exported = await exportDocx(docxBytes, { DOMParserImpl: DOMParser, filename: "plain-paragraphs" });

    let response = exported.markdown;
    response = response.replace("the first paragraph", "the {++truly ++}first paragraph");
    response = response.replace("This is the second", "{--This is the second--}");

    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);

    const state = createRatificationState(result.edits);
    // Reject the deletion, keep the insertion -- exercises acceptedEdits() actually
    // filtering, not just passing every resolved edit straight through.
    const delRow = state.rows.find((r) => r.edit.type === "del");
    state.setDecision(delRow.id, "reject");

    const rewritten = await buildReviewedDocx({
      docxBytes,
      acceptedEdits: state.acceptedEdits(),
      sourceMap: exported.sourceMap,
      author: "AutoReviewer — Test Persona",
      date: "2026-01-01T00:00:00Z",
      DOMParserImpl: DOMParser,
      XMLSerializerImpl: XMLSerializer,
    });

    expect(rewritten).toBeInstanceOf(ArrayBuffer);
    const zip = await unzip(rewritten);
    const docXml = await readEntry(zip, "word/document.xml");
    expect(docXml).toContain('<w:ins w:id="1" w:author="AutoReviewer — Test Persona"');
    expect(docXml).toContain("truly ");
    // The rejected deletion never reached injectEdits -- no w:del anywhere in the output.
    expect(docXml).not.toContain("<w:del ");

    // Re-exporting the result should show the insertion as a genuinely TRACKED change
    // (there's no "accepted" bit in OOXML -- ratifying it here only decided which edits
    // reached injectEdits at all; it's still {++...++} until a human accepts it in Word).
    const reExported = await exportDocx(rewritten, { DOMParserImpl: DOMParser, filename: "plain-paragraphs" });
    expect(reExported.markdown).toContain("This is the {++truly ++}first paragraph");
  });

  it("wires new comments through upsertComments -- a document with no prior comments gets a fresh comments.xml", async () => {
    const docxBytes = loadDocxBytes("plain-paragraphs");
    const exported = await exportDocx(docxBytes, { DOMParserImpl: DOMParser, filename: "plain-paragraphs" });

    const response = exported.markdown.replace(
      "nothing tracked.",
      "nothing tracked.{>>Please double-check this framing.<<}"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);

    const state = createRatificationState(result.edits);
    const rewritten = await buildReviewedDocx({
      docxBytes,
      acceptedEdits: state.acceptedEdits(),
      sourceMap: exported.sourceMap,
      author: "AutoReviewer — Test Persona",
      date: "2026-01-01T00:00:00Z",
      DOMParserImpl: DOMParser,
      XMLSerializerImpl: XMLSerializer,
    });

    const zip = await unzip(rewritten);
    expect(zip.order).toContain("word/comments.xml");
    const commentsXml = await readEntry(zip, "word/comments.xml");
    expect(commentsXml).toContain("Please double-check this framing.");
    expect(commentsXml).toContain('w:author="AutoReviewer — Test Persona"');
  });
});

// M4b's INJECTED-step audit wiring (app.js assembles buildAuditRecord's `details` from the
// same real pieces buildReviewedDocx above uses -- docxBytes, sourceMap, ratify rows, and
// the actual written output bytes). Exercised here against real fixtures + real
// crypto.subtle (no digestImpl override), in the default Node environment, rather than via
// a happy-dom click: buildReviewedDocx crosses real Streams-API boundaries the Inject
// button's happy-dom test above already opts out of (see this file's header comment).
describe("app.js's audit wiring: buildAuditRecord over the real pieces the Inject handler assembles", () => {
  it("produces a spec-complete record whose source.sha256 matches the real sourceMap.docHash", async () => {
    const docxBytes = loadDocxBytes("plain-paragraphs");
    const exported = await exportDocx(docxBytes, { DOMParserImpl: DOMParser, filename: "plain-paragraphs" });

    let response = exported.markdown;
    response = response.replace("the first paragraph", "the {++truly ++}first paragraph");
    response = response.replace("This is the second", "{--This is the second--}");

    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);

    const ratify = createRatificationState(result.edits);
    const delRow = ratify.rows.find((r) => r.edit.type === "del");
    ratify.setDecision(delRow.id, "reject");

    const author = "AutoReviewer — Test Persona";
    const rewritten = await buildReviewedDocx({
      docxBytes,
      acceptedEdits: ratify.acceptedEdits(),
      sourceMap: exported.sourceMap,
      author,
      date: "2026-01-01T00:00:00Z",
      DOMParserImpl: DOMParser,
      XMLSerializerImpl: XMLSerializer,
    });

    const record = await buildAuditRecord({
      promptVersion: "m4a-2026.07-1",
      timestamps: { loaded: "2026-01-01T00:00:00Z", injected: "2026-01-01T00:01:00Z" },
      filename: "plain-paragraphs",
      docxBytes,
      outputBytes: rewritten,
      response,
      sourceMap: exported.sourceMap,
      persona: null,
      validationAttempts: [{ ts: "2026-01-01T00:00:30Z", result: "ok" }],
      rows: ratify.rows,
      author,
    });

    expect(record.source.sha256).toBe(exported.sourceMap.docHash);
    expect(record.persona).toEqual({ name: "Default Persona (built-in)" });
    expect(record.edits).toHaveLength(2);
    expect(record.edits.find((e) => e.type === "ins").decision).toBe("accept");
    expect(record.edits.find((e) => e.type === "del").decision).toBe("reject");
    // The rejected deletion never reached injectEdits (see the test above) -- counts
    // reflect only what actually got injected.
    expect(record.injection.counts).toEqual({ ins: 1, del: 0, sub: 0, comment: 0 });
  });
});
