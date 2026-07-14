// Step 3 of the reviewer-pass-slicer task: slice rendering on top of clusterPasses.
import { describe, expect, it } from "vitest";
import { DOMParser } from "@xmldom/xmldom";
import { extractObservations } from "../src/ooxml/observations.js";
import { clusterPasses } from "../src/passes.js";
import { renderSlice, sliceFilename } from "../src/ooxml/slice.js";
import { buildAuditFixtureDocx } from "./helpers/auditFixture.js";
import { unzip } from "../src/zip/reader.js";
import { writeZip } from "../src/zip/writer.js";
import { loadDocxBytes } from "./helpers/docx.js";

async function auditPasses() {
  const bytes = await buildAuditFixtureDocx();
  const { observations } = await extractObservations(bytes, { DOMParserImpl: DOMParser, filename: "audit-fixture" });
  const { passes } = clusterPasses(observations);
  return { bytes, passes };
}

function findPass(passes, author, passDate) {
  return passes.find((p) => p.author === author && p.passDate === passDate);
}

describe("renderSlice: Jim Smith, pass 1 (2026-05-14)", () => {
  it("renders his own insertion and comment live; his own pass-2 deletion and Katie's insertion as accepted plain text", async () => {
    const { bytes, passes } = await auditPasses();
    const pass = findPass(passes, "Jim Smith", "2026-05-14");
    const md = await renderSlice(bytes, pass, { DOMParserImpl: DOMParser, filename: "audit-fixture", generated: "2026-07-14" });

    expect(md).toContain("{++all covered carriers++}");
    expect(md).toContain(
      '{==This sentence has a discussion thread attached to it.==}{>>Jim Smith [2026-05-14]: Is this the right threshold?<<}'
    );

    // Pass 2's own deletion is not this pass -- accepted view removes it entirely.
    expect(md).not.toContain("should be removed entirely");
    expect(md).not.toContain("{--");
    expect(md).toContain("This clause is obsolete and .");

    // Katie's insertion belongs to a different reviewer entirely -- accepted view, plain.
    expect(md).toContain("The deadline is 60 days after publication.");
    expect(md).not.toContain("{++60 days++}");

    // Katie's reply to this very comment is a different pass -- not shown in Jim's own slice.
    expect(md).not.toContain("Agreed");
  });

  it("frontmatter matches the pass", async () => {
    const { bytes, passes } = await auditPasses();
    const pass = findPass(passes, "Jim Smith", "2026-05-14");
    const md = await renderSlice(bytes, pass, { DOMParserImpl: DOMParser, filename: "audit-fixture", generated: "2026-07-14" });
    const frontmatter = md.split("\n\n")[0];
    expect(frontmatter).toBe(
      [
        "---",
        "doc: audit-fixture.docx",
        "reviewer: Jim Smith",
        "pass: 2026-05-14",
        "pass_window: 2026-05-14T09:12Z – 2026-05-14T09:30Z",
        "edits: 1",
        "comments: 1",
        "replies: 0",
        "generated: 2026-07-14",
        "---",
      ].join("\n")
    );
  });
});

describe("renderSlice: Katie Chen (2026-05-15)", () => {
  it("embeds Jim's comment read-only as context for her reply, and renders her own insertion/point comment live", async () => {
    const { bytes, passes } = await auditPasses();
    const pass = findPass(passes, "Katie Chen", "2026-05-15");
    const md = await renderSlice(bytes, pass, { DOMParserImpl: DOMParser, filename: "audit-fixture", generated: "2026-07-14" });

    expect(md).toContain("{++60 days++}");

    // Jim's root comment (a different reviewer, different pass) shown read-only for
    // context, immediately followed by Katie's own reply -- never orphaned.
    expect(md).toContain(
      '{==This sentence has a discussion thread attached to it.==}' +
        '{>>Jim Smith [2026-05-14]: Is this the right threshold?<<}' +
        '{>>↳ reply to Jim Smith — Katie Chen [2026-05-15]: Agreed -- flagged for legal review.<<}'
    );

    // Her own point comment: {==¶==} placeholder + containing-sentence context.
    expect(md).toContain(
      '{==¶==}{>>Katie Chen [2026-05-15] (context: "The deadline sits within this second sentence for testing."):' +
        " Does this point still hold after the revision above?<<}"
    );

    // Jim's own edits (both passes) belong to a different reviewer -- accepted view.
    expect(md).toContain("The rule shall apply to all covered carriers.");
    expect(md).not.toContain("{++all covered carriers++}");
    expect(md).not.toContain("should be removed entirely");

    // Jim's other, unrelated resolved comment isn't connected to Katie's material at all.
    expect(md).not.toContain("Please cite the authority here");
  });
});

describe("sliceFilename", () => {
  it("slugifies doc/author and uses the pass date", async () => {
    const { passes } = await auditPasses();
    const pass = findPass(passes, "Jim Smith", "2026-05-16");
    expect(sliceFilename("Air Cargo RIA v2", pass)).toBe("air-cargo-ria-v2_jim-smith_2026-05-16.md");
  });

  it("uses 'undated' for an undated pass", () => {
    const undatedPass = { author: "Al", undated: true, passDate: null };
    expect(sliceFilename("doc", undatedPass)).toBe("doc_al_undated.md");
  });
});

// A second, minimal fixture -- two authors, each with their own single-author
// substitution -- since the audit fixture has no {~~sub~~} case at all. Verifies a sub
// belonging to the target pass renders as one substitution token, and a sub belonging to
// someone else renders as its accepted (replacement-text-only) view.
const TWO_AUTHOR_SUB_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
<w:body>
<w:p><w:r><w:t xml:space="preserve">The deadline is </w:t></w:r><w:del w:id="1" w:author="Al" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>30 days</w:delText></w:r></w:del><w:ins w:id="2" w:author="Al" w:date="2026-01-01T00:00:00Z"><w:r><w:t>60 days</w:t></w:r></w:ins><w:r><w:t xml:space="preserve"> after publication.</w:t></w:r></w:p>
<w:p><w:r><w:t xml:space="preserve">The meeting is on </w:t></w:r><w:del w:id="3" w:author="Bo" w:date="2026-01-02T00:00:00Z"><w:r><w:delText>Monday</w:delText></w:r></w:del><w:ins w:id="4" w:author="Bo" w:date="2026-01-02T00:00:00Z"><w:r><w:t>Tuesday</w:t></w:r></w:ins><w:r><w:t>.</w:t></w:r></w:p>
<w:sectPr/>
</w:body>
</w:document>`;

async function buildTwoAuthorSubBytes() {
  const zip = await unzip(loadDocxBytes("tracked-changes"));
  return writeZip(zip, { "word/document.xml": TWO_AUTHOR_SUB_DOCUMENT_XML });
}

describe("renderSlice: substitutions", () => {
  it("renders the target pass's own substitution as {~~old~>new~~} and another reviewer's as accepted replacement text", async () => {
    const bytes = await buildTwoAuthorSubBytes();
    const { observations } = await extractObservations(bytes, { DOMParserImpl: DOMParser, filename: "two-author-sub" });
    const { passes } = clusterPasses(observations);
    const al = passes.find((p) => p.author === "Al");

    const md = await renderSlice(bytes, al, { DOMParserImpl: DOMParser, filename: "two-author-sub", generated: "2026-07-14" });
    expect(md).toContain("{~~30 days~>60 days~~}");
    expect(md).toContain("The meeting is on Tuesday.");
    expect(md).not.toContain("Monday");
    expect(md).not.toContain("{~~Monday~>Tuesday~~}");
  });
});
