// M4d PR-2: two new load-time preflight fences (docs/plans/m4d... F-2, F-3).
//
// Structural fence (permanent, until child-aware splitRun lands): a body run containing
// w:br, w:cr, w:tab, or w:delText crashes inject.js's splitRun later with "invalid range"
// (F-2 -- export.js and inject.js disagreed about a run's plain-text length). No committed
// fixture has one of these inside a paragraph, so each trigger is synthesized here by
// swapping a hand-written word/document.xml into a real fixture's ZIP shell, the same
// technique tests/app.chunkFlow.test.js uses.
//
// Annotation fence (temporary -- first thing M6a removes): ANY pre-existing comment or
// tracked change, not just the comment-REPLY case US-7 already blocked. comments-threaded
// .docx already has a top-level (non-reply) comment, so it doubles as that fixture.
import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { describe, expect, it } from "vitest";
import { checkPreflight, loadDocxFromBytes, STRUCTURAL_FENCE_MESSAGE, ANNOTATION_FENCE_MESSAGE, D4_ERROR_MESSAGE } from "../src/ui/load.js";
import { unzip, readEntry } from "../src/zip/reader.js";
import { writeZip } from "../src/zip/writer.js";
import { DOMParser, loadDocxBytes } from "./helpers/docx.js";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const fixturesDir = path.join(root, "fixtures");

async function docxWithBody(bodyInner) {
  const buf = readFileSync(path.join(fixturesDir, "plain-paragraphs.docx"));
  const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
  const zip = await unzip(ab);
  const orig = await readEntry(zip, "word/document.xml");
  const rootOpen = orig.slice(0, orig.indexOf(">", orig.indexOf("<w:document")) + 1);
  const tail = orig.slice(orig.indexOf("<w:sectPr")); // "<w:sectPr...>...</w:sectPr></w:body></w:document>"
  const newXml = rootOpen + "<w:body>" + bodyInner + tail;
  return writeZip(zip, { "word/document.xml": newXml });
}

describe("checkPreflight: M4d structural + annotation fences (unit)", () => {
  it("blocks with the structural message when structuralHazard is set", () => {
    const result = checkPreflight({ counts: { ins: 0, del: 0, sub: 0 }, comments: {}, structuralHazard: true });
    expect(result.ok).toBe(false);
    expect(result.message).toContain(STRUCTURAL_FENCE_MESSAGE);
  });

  it("blocks with the annotation message on ANY comment, not just a reply", () => {
    const result = checkPreflight({
      counts: { ins: 0, del: 0, sub: 0 },
      comments: { c1: { id: "c1", parentId: null } }, // top-level, not a reply
      structuralHazard: false,
    });
    expect(result.ok).toBe(false);
    expect(result.message).toContain(ANNOTATION_FENCE_MESSAGE);
  });

  it("reports every applicable reason together, not just the first hit", () => {
    const result = checkPreflight({
      counts: { ins: 1, del: 0, sub: 0 },
      comments: { c1: { id: "c1", parentId: null } },
      structuralHazard: true,
    });
    expect(result.ok).toBe(false);
    expect(result.message).toContain(STRUCTURAL_FENCE_MESSAGE);
    expect(result.message).toContain(D4_ERROR_MESSAGE);
    expect(result.message).toContain(ANNOTATION_FENCE_MESSAGE);
  });
});

describe("loadDocxFromBytes: structural fence triggers (end-to-end, synthesized fixtures)", () => {
  it("rejects a soft line break (w:br) inside a paragraph", async () => {
    const bytes = await docxWithBody(
      '<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>After</w:t></w:r></w:p>'
    );
    const result = await loadDocxFromBytes(bytes, { originalFilename: "hazard-br.docx", DOMParserImpl: DOMParser });
    expect(result.ok).toBe(false);
    expect(result.message).toContain(STRUCTURAL_FENCE_MESSAGE);
  });

  it("rejects a manual page break / carriage return (w:cr) inside a paragraph", async () => {
    const bytes = await docxWithBody('<w:p><w:r><w:t>Before</w:t></w:r><w:r><w:cr/></w:r><w:r><w:t>After</w:t></w:r></w:p>');
    const result = await loadDocxFromBytes(bytes, { originalFilename: "hazard-cr.docx", DOMParserImpl: DOMParser });
    expect(result.ok).toBe(false);
    expect(result.message).toContain(STRUCTURAL_FENCE_MESSAGE);
  });

  it("rejects a tab (w:tab) inside a paragraph", async () => {
    const bytes = await docxWithBody('<w:p><w:r><w:t>Before</w:t><w:tab/><w:t>After</w:t></w:r></w:p>');
    const result = await loadDocxFromBytes(bytes, { originalFilename: "hazard-tab.docx", DOMParserImpl: DOMParser });
    expect(result.ok).toBe(false);
    expect(result.message).toContain(STRUCTURAL_FENCE_MESSAGE);
  });

  it("rejects a tracked deletion's w:delText inside a paragraph", async () => {
    const bytes = await docxWithBody(
      '<w:p><w:r><w:t>Before </w:t></w:r><w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">' +
        '<w:r><w:delText>gone</w:delText></w:r></w:del><w:r><w:t> after</w:t></w:r></w:p>'
    );
    const result = await loadDocxFromBytes(bytes, { originalFilename: "hazard-deltext.docx", DOMParserImpl: DOMParser });
    expect(result.ok).toBe(false);
    // Also a pre-existing tracked change, so both fences fire -- this fixture doubles as
    // proof the two messages coexist rather than one masking the other.
    expect(result.message).toContain(STRUCTURAL_FENCE_MESSAGE);
    expect(result.message).toContain(D4_ERROR_MESSAGE);
    expect(result.message).toContain(ANNOTATION_FENCE_MESSAGE);
  });

  it("loads a clean document with none of the hazards", async () => {
    const bytes = await docxWithBody("<w:p><w:r><w:t>Clean paragraph, no hazards.</w:t></w:r></w:p>");
    const result = await loadDocxFromBytes(bytes, { originalFilename: "clean.docx", DOMParserImpl: DOMParser });
    expect(result.ok).toBe(true);
  });
});

describe("loadDocxFromBytes: annotation fence triggers (end-to-end, real fixtures)", () => {
  it("rejects a document with pre-existing tracked insertions/deletions", async () => {
    const bytes = loadDocxBytes("tracked-changes");
    const result = await loadDocxFromBytes(bytes, { originalFilename: "tracked-changes.docx", DOMParserImpl: DOMParser });
    expect(result.ok).toBe(false);
    expect(result.message).toContain(ANNOTATION_FENCE_MESSAGE);
  });

  it("rejects a document with pre-existing comments, even non-reply ones", async () => {
    const bytes = loadDocxBytes("comments-threaded");
    const result = await loadDocxFromBytes(bytes, { originalFilename: "comments-threaded.docx", DOMParserImpl: DOMParser });
    expect(result.ok).toBe(false);
    expect(result.message).toContain(ANNOTATION_FENCE_MESSAGE);
  });
});
