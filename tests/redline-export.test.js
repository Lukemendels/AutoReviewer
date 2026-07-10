// Redline export shell (M6 first slice). Asserts against exportDocx directly, per the task
// spec ("Do not test the DOM shell") -- src/ui/redline-export.js is a thin DOM adapter over
// the same exportDocx() src/ooxml/export.js already exercises in tests/export.test.js.
//
// No committed fixture has both pre-existing tracked changes AND a comment reply (the reply
// case is deliberately excluded from the corpus -- see src/ui/load.js's COMMENT_REPLY_MESSAGE
// -- since checkPreflight blocks it in every other flow). This module's whole reason to exist
// is exporting exactly that document, so the fixture is synthesized here: start from
// comments-threaded.docx's container (it already wires word/commentsExtended.xml's relationship
// + content-type override) and swap in hand-written document.xml/comments.xml/
// commentsExtended.xml via writeZip, the same technique tests/app.chunkFlow.test.js uses for
// its oversized multi-heading document.
import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { exportDocx } from "../src/ooxml/export.js";
import { unzip, readEntry } from "../src/zip/reader.js";
import { writeZip } from "../src/zip/writer.js";
import { loadDocxBytes } from "./helpers/docx.js";

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t xml:space="preserve">The rule shall apply to </w:t></w:r><w:ins w:id="10" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>all covered carriers</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">This clause is obsolete and </w:t></w:r><w:del w:id="11" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>should be removed entirely</w:delText></w:r></w:del><w:r><w:t>.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">The deadline is </w:t></w:r><w:del w:id="12" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30 days</w:delText></w:r></w:del><w:ins w:id="13" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>60 days</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> after publication.</w:t></w:r></w:p>
<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>This sentence has a discussion thread attached to it.</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>This sentence has a resolved comment attached to it.</w:t></w:r><w:commentRangeEnd w:id="2"/><w:r><w:commentReference w:id="2"/></w:r></w:p>
<w:sectPr/>
</w:body>
</w:document>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
<w:comment w:id="0" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:p w14:paraId="10000000"><w:r><w:t>Is this the right threshold?</w:t></w:r></w:p></w:comment>
<w:comment w:id="1" w:author="Reviewer B" w:date="2026-01-02T00:00:00Z"><w:p w14:paraId="10000001"><w:r><w:t>Agreed -- flagged for legal review.</w:t></w:r></w:p></w:comment>
<w:comment w:id="2" w:author="Reviewer C" w:date="2026-01-03T00:00:00Z"><w:p w14:paraId="10000002"><w:r><w:t>Please cite the authority here.</w:t></w:r></w:p></w:comment>
</w:comments>`;

const COMMENTS_EXTENDED_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
<w15:commentEx w15:paraId="10000000" w15:done="0"/>
<w15:commentEx w15:paraId="10000001" w15:paraIdParent="10000000" w15:done="0"/>
<w15:commentEx w15:paraId="10000002" w15:done="1"/>
</w15:commentsEx>`;

async function buildSyntheticDocxBytes() {
  const zip = await unzip(loadDocxBytes("comments-threaded"));
  return writeZip(zip, {
    "word/document.xml": DOCUMENT_XML,
    "word/comments.xml": COMMENTS_XML,
    "word/commentsExtended.xml": COMMENTS_EXTENDED_XML,
  });
}

async function exportSynthetic() {
  const bytes = await buildSyntheticDocxBytes();
  return exportDocx(bytes, { DOMParserImpl: DOMParser, annotate: true, filename: "synthetic" });
}

describe("redline export (M6 first slice): tracked changes + comment reply + resolved comment", () => {
  it("does not throw and does not invoke any preflight check", async () => {
    // A regular load (src/ui/load.js) would reject this exact document at the door: it has
    // pre-existing tracked changes (D4_ERROR_MESSAGE) AND a comment reply
    // (COMMENT_REPLY_MESSAGE). exportDocx succeeding here proves this module's whole premise:
    // it drives exportDocx directly and never routes through checkPreflight.
    await expect(exportSynthetic()).resolves.toBeTruthy();
  });

  it("renders a tracked insertion as {++...++}", async () => {
    const { markdown, counts } = await exportSynthetic();
    expect(markdown).toContain("{++all covered carriers++}");
    expect(counts.ins).toBe(1);
  });

  it("renders a tracked deletion as {--...--}", async () => {
    const { markdown, counts } = await exportSynthetic();
    expect(markdown).toContain("{--should be removed entirely--}");
    expect(counts.del).toBe(1);
  });

  it("renders adjacent deletion+insertion as a single {~~old~>new~~} substitution", async () => {
    const { markdown, counts } = await exportSynthetic();
    expect(markdown).toContain("{~~30 days~>60 days~~}");
    expect(counts.sub).toBe(1);
  });

  it("renders an anchored comment as {==text==}{>>Author (date): text<<}", async () => {
    const { markdown } = await exportSynthetic();
    expect(markdown).toContain("{==This sentence has a discussion thread attached to it.==}");
    expect(markdown).toContain("{>>Reviewer A (2026-01-01): Is this the right threshold?<<}");
  });

  it("renders a comment reply inside the same thread with a depth marker, parent not duplicated", async () => {
    const { markdown } = await exportSynthetic();
    expect(markdown).toContain("↳ Reviewer B (2026-01-02): Agreed -- flagged for legal review.");
    // Exactly one rendering of the parent thread's anchor -- the reply must not cause the
    // parent comment to be emitted a second time.
    const parentOccurrences = markdown.split("Is this the right threshold?").length - 1;
    expect(parentOccurrences).toBe(1);
  });

  it("carries the [resolved] marker on a resolved comment", async () => {
    const { markdown } = await exportSynthetic();
    expect(markdown).toContain("{>>Reviewer C (2026-01-03) [resolved]: Please cite the authority here.<<}");
  });

  it("reports comment counts distinguishing top-level threads from total nodes", async () => {
    const { comments } = await exportSynthetic();
    const nodes = Object.values(comments);
    expect(nodes).toHaveLength(3);
    expect(nodes.filter((c) => c.parentId == null)).toHaveLength(2);
  });
});
