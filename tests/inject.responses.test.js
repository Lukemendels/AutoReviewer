import { describe, expect, it } from "vitest";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { parseXml } from "../src/ooxml/parse.js";
import { injectResponses } from "../src/ooxml/inject.js";

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>This is some text with </w:t></w:r>
      <w:ins w:id="10" w:author="Reviewer A" w:date="2026-07-02T00:00:00Z">
        <w:r><w:t>an inserted change</w:t></w:r>
      </w:ins>
      <w:r><w:t> and more text.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

describe("injectResponses (M6c)", () => {
  it("injects a threaded reply and resolves the comment thread", () => {
    const doc = parseXml(DOCUMENT_XML, DOMParser);
    const sourceMap = {
      annotations: {
        C1: { type: "comment", id: "5" },
      },
    };

    const decisions = [
      {
        label: "C1",
        type: "comment",
        originalId: "5",
        reply: "Agreed -- resolved.",
        resolve: true,
      },
    ];

    const result = injectResponses(doc, decisions, sourceMap, {
      author: "AutoReviewer — Test",
      date: "2026-07-15T00:00:00Z",
    });

    expect(result.newComments).toHaveLength(1);
    expect(result.newComments[0]).toEqual({
      id: expect.any(Number),
      parentId: "5",
      author: "AutoReviewer — Test",
      date: "2026-07-15T00:00:00Z",
      text: "Agreed -- resolved.",
      done: true,
    });
  });

  it("injects a recommendation point comment adjacent to a revision element", () => {
    const doc = parseXml(DOCUMENT_XML, DOMParser);
    const sourceMap = {
      annotations: {
        R1: { type: "revision", id: "10" },
      },
    };

    const decisions = [
      {
        label: "R1",
        type: "revision",
        originalId: "10",
        decision: "accept",
        rationale: "Aligns with OMB guidelines.",
      },
    ];

    const result = injectResponses(doc, decisions, sourceMap, {
      author: "AutoReviewer — Test",
      date: "2026-07-15T00:00:00Z",
    });

    expect(result.newComments).toHaveLength(1);
    expect(result.newComments[0]).toEqual({
      id: expect.any(Number),
      author: "AutoReviewer — Test",
      date: "2026-07-15T00:00:00Z",
      text: "[AR:accept] Aligns with OMB guidelines.",
    });

    const ins = doc.getElementsByTagName("w:ins")[0];
    expect(ins).toBeTruthy();

    const sibling1 = ins.nextSibling;
    expect(sibling1.localName).toBe("commentRangeStart");
    expect(sibling1.getAttribute("w:id")).toBe(String(result.newComments[0].id));

    const sibling2 = sibling1.nextSibling;
    expect(sibling2.localName).toBe("commentRangeEnd");
    expect(sibling2.getAttribute("w:id")).toBe(String(result.newComments[0].id));

    const sibling3 = sibling2.nextSibling;
    expect(sibling3.localName).toBe("r");
    const ref = sibling3.getElementsByTagName("w:commentReference")[0];
    expect(ref).toBeTruthy();
    expect(ref.getAttribute("w:id")).toBe(String(result.newComments[0].id));
  });

  // M6 is additive-only: the Respond injection path must never modify existing revisions,
  // only add comment markers adjacent to them. This is the missing guard flagged in the M6
  // code review -- the code appeared correct, but nothing pinned it down.
  it("is additive-only: existing w:ins/w:del elements are byte-identical after injecting a threaded reply and a point recommendation", () => {
    const documentXmlWithDeletion = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>This is some text with </w:t></w:r>
      <w:ins w:id="10" w:author="Reviewer A" w:date="2026-07-02T00:00:00Z">
        <w:r><w:t>an inserted change</w:t></w:r>
      </w:ins>
      <w:r><w:t> and </w:t></w:r>
      <w:del w:id="11" w:author="Reviewer A" w:date="2026-07-02T00:00:00Z">
        <w:r><w:delText>a deleted phrase</w:delText></w:r>
      </w:del>
      <w:r><w:t> and more text.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

    const doc = parseXml(documentXmlWithDeletion, DOMParser);
    const serializer = new XMLSerializer();
    const revisionEl = (el) => el.localName === "ins" || el.localName === "del";

    const before = [...doc.getElementsByTagName("*")].filter(revisionEl).map((el) => serializer.serializeToString(el));
    expect(before).toHaveLength(2);

    const sourceMap = {
      annotations: {
        C1: { type: "comment", id: "5" },
        R1: { type: "revision", id: "10" },
      },
    };
    const decisions = [
      { label: "C1", type: "comment", originalId: "5", reply: "Agreed -- resolved.", resolve: true },
      { label: "R1", type: "revision", originalId: "10", decision: "accept", rationale: "Aligns with OMB guidelines." },
    ];

    injectResponses(doc, decisions, sourceMap, {
      author: "AutoReviewer — Test",
      date: "2026-07-15T00:00:00Z",
    });

    const after = [...doc.getElementsByTagName("*")].filter(revisionEl).map((el) => serializer.serializeToString(el));
    expect(after).toHaveLength(before.length);
    before.forEach((xml, i) => expect(after[i]).toBe(xml));
  });
});
