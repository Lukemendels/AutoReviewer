// ooxml/comments.js (spec §9.2; M3b plan's D6 smaller-cut scope). Written and red before
// upsertComments has a real implementation.
import { describe, expect, it } from "vitest";
import { DOMParser } from "@xmldom/xmldom";
import { loadDocxBytes } from "./helpers/docx.js";
import { unzip, readEntry } from "../src/zip/reader.js";
import { parseXml } from "../src/ooxml/parse.js";
import { buildCommentsData } from "../src/ooxml/export.js";
import { upsertComments } from "../src/ooxml/comments.js";

const OPTS = { DOMParserImpl: DOMParser, XMLSerializerImpl: (await import("@xmldom/xmldom")).XMLSerializer };

async function loadRealCommentParts(fixtureName) {
  const zip = await unzip(loadDocxBytes(fixtureName));
  return {
    commentsXml: await readEntry(zip, "word/comments.xml"),
    commentsExtendedXml: await readEntry(zip, "word/commentsExtended.xml"),
    relsXml: await readEntry(zip, "word/_rels/document.xml.rels"),
    contentTypesXml: await readEntry(zip, "[Content_Types].xml"),
  };
}

function paraIdsIn(xmlString) {
  const doc = parseXml(xmlString, DOMParser);
  const ids = [];
  for (const el of doc.getElementsByTagName("*")) {
    if (el.localName === "p") {
      const pid = el.getAttributeNS("http://schemas.microsoft.com/office/word/2010/wordml", "paraId");
      if (pid) ids.push(pid);
    }
  }
  return ids;
}
function commentExesIn(xmlString) {
  const doc = parseXml(xmlString, DOMParser);
  return [...doc.getElementsByTagName("*")].filter((el) => el.localName === "commentEx");
}

describe("upsertComments: create-from-scratch", () => {
  it("creates all four parts when none exist, with valid XML declarations and namespaces", () => {
    const result = upsertComments(
      {},
      [{ id: 0, author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z", text: "Please verify this framing." }],
      OPTS
    );

    expect(result.commentsXml).toMatch(/^<\?xml version="1\.0" encoding="UTF-8"/);
    expect(result.commentsXml).toContain('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"');
    expect(result.commentsExtendedXml).toMatch(/^<\?xml version="1\.0" encoding="UTF-8"/);
    expect(result.commentsExtendedXml).toContain('xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"');

    const commentsDoc = parseXml(result.commentsXml, DOMParser);
    const comments = [...commentsDoc.getElementsByTagName("*")].filter((el) => el.localName === "comment");
    expect(comments).toHaveLength(1);
    expect(comments[0].getAttribute("w:id")).toBe("0");
    expect(comments[0].getAttribute("w:author")).toBe("AutoReviewer — Test");
    expect(comments[0].textContent).toContain("Please verify this framing.");

    const paraIds = paraIdsIn(result.commentsXml);
    expect(paraIds).toHaveLength(1);
    expect(paraIds[0]).toMatch(/^[0-9A-F]{8}$/);

    const exes = commentExesIn(result.commentsExtendedXml);
    expect(exes).toHaveLength(1);
    expect(exes[0].getAttribute("w15:paraId")).toBe(paraIds[0]);
    expect(exes[0].getAttribute("w15:done")).toBe("0");
  });

  it("wires the document rels entries and [Content_Types].xml overrides for a fresh document", () => {
    // A minimal but realistic starting point -- a real docx always already has rels/
    // content-types (unlike comments.xml, which may genuinely be absent).
    const result = upsertComments(
      {
        relsXml:
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',
        contentTypesXml:
          '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
      },
      [{ id: 0, author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z", text: "A note." }],
      OPTS
    );

    expect(result.relsXml).toContain("http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments");
    expect(result.relsXml).toContain('Target="comments.xml"');
    expect(result.relsXml).toContain("http://schemas.microsoft.com/office/2011/relationships/commentsExtended");
    expect(result.relsXml).toContain('Target="commentsExtended.xml"');
    // The pre-existing relationship is untouched, not clobbered.
    expect(result.relsXml).toContain('Target="styles.xml"');

    expect(result.contentTypesXml).toContain('PartName="/word/comments.xml"');
    expect(result.contentTypesXml).toContain("wordprocessingml.comments+xml");
    expect(result.contentTypesXml).toContain('PartName="/word/commentsExtended.xml"');
    expect(result.contentTypesXml).toContain("wordprocessingml.commentsExtended+xml");
    expect(result.contentTypesXml).toContain('PartName="/word/document.xml"');
  });

  it("does not implement commentsIds.xml or commentsExtensible.xml (D6 scope guard)", () => {
    const result = upsertComments({}, [{ id: 0, author: "A", date: "2026-01-01T00:00:00Z", text: "x" }], OPTS);
    expect(result.commentsIdsXml).toBeFalsy();
    expect(result.commentsExtensibleXml).toBeFalsy();
  });
});

describe("upsertComments: extend-existing", () => {
  it("appends a new comment to real Word-authored comment parts without disturbing the existing ones", async () => {
    const existing = await loadRealCommentParts("comments-threaded");
    const existingCommentCount = [...parseXml(existing.commentsXml, DOMParser).getElementsByTagName("*")].filter(
      (el) => el.localName === "comment"
    ).length;
    const existingExCount = commentExesIn(existing.commentsExtendedXml).length;

    const result = upsertComments(
      existing,
      [{ id: 3, author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z", text: "A freshly injected comment." }],
      OPTS
    );

    const commentsDoc = parseXml(result.commentsXml, DOMParser);
    const comments = [...commentsDoc.getElementsByTagName("*")].filter((el) => el.localName === "comment");
    expect(comments).toHaveLength(existingCommentCount + 1);
    expect(comments.map((c) => c.getAttribute("w:id"))).toContain("3");
    // Pre-existing comment text survives untouched.
    expect(result.commentsXml).toContain("Is this the right threshold?");
    expect(result.commentsXml).toContain("Please cite the authority here.");

    expect(commentExesIn(result.commentsExtendedXml)).toHaveLength(existingExCount + 1);

    // Rels/content-types already had the linkage -- extending must not duplicate it.
    const relCount = (result.relsXml.match(/relationships\/comments"/g) || []).length;
    expect(relCount).toBe(1);
    const ctCount = (result.contentTypesXml.match(/wordprocessingml\.comments\+xml/g) || []).length;
    expect(ctCount).toBe(1);
  });

  it("generates a fresh paraId that doesn't collide with any existing paraId in the document", async () => {
    const existing = await loadRealCommentParts("comments-threaded");
    const before = new Set(paraIdsIn(existing.commentsXml));

    const result = upsertComments(existing, [{ id: 3, author: "A", date: "2026-01-01T00:00:00Z", text: "x" }], OPTS);

    const after = paraIdsIn(result.commentsXml);
    expect(after).toHaveLength(before.size + 1);
    const newParaId = after.find((id) => !before.has(id));
    expect(newParaId).toBeTruthy();
    expect(newParaId).toMatch(/^[0-9A-F]{8}$/);
  });
});

describe("upsertComments: threading (parentId)", () => {
  it("sets w15:paraIdParent when a new comment replies to an existing comment", async () => {
    const existing = await loadRealCommentParts("comments-threaded");
    // comments-threaded.docx's first comment has w:id "0".
    const result = upsertComments(
      existing,
      [{ id: 3, author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z", text: "A threaded reply.", parentId: 0 }],
      OPTS
    );

    const parentParaId = paraIdsIn(existing.commentsXml)[0];
    const newExEl = commentExesIn(result.commentsExtendedXml).find((el) => !commentExesIn(existing.commentsExtendedXml).some((old) => old.getAttribute("w15:paraId") === el.getAttribute("w15:paraId")));
    expect(newExEl).toBeTruthy();
    expect(newExEl.getAttribute("w15:paraIdParent")).toBe(parentParaId);
  });

  it("threads a reply to a comment created in the SAME batch", () => {
    const result = upsertComments(
      {},
      [
        { id: 0, author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z", text: "Top-level comment." },
        { id: 1, author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z", text: "A reply.", parentId: 0 },
      ],
      OPTS
    );
    const exes = commentExesIn(result.commentsExtendedXml);
    expect(exes).toHaveLength(2);
    const parentParaId = exes[0].getAttribute("w15:paraId");
    expect(exes[1].getAttribute("w15:paraIdParent")).toBe(parentParaId);
    expect(exes[0].getAttribute("w15:paraIdParent")).toBeFalsy();
  });
});

describe("upsertComments: round-trips through the reader", () => {
  it("a newly-injected comment is itself re-exportable/readable via export.js's buildCommentsData", () => {
    const result = upsertComments(
      {},
      [
        { id: 0, author: "AutoReviewer — Persona", date: "2026-01-01T00:00:00Z", text: "Top-level note." },
        { id: 1, author: "AutoReviewer — Persona", date: "2026-01-02T00:00:00Z", text: "A reply.", parentId: 0 },
      ],
      OPTS
    );

    const { comments, childrenMap } = buildCommentsData(result.commentsXml, result.commentsExtendedXml, DOMParser);
    expect(Object.keys(comments)).toHaveLength(2);
    expect(comments["0"].author).toBe("AutoReviewer — Persona");
    expect(comments["0"].text).toBe("Top-level note.");
    expect(comments["0"].done).toBe(false);
    expect(comments["1"].text).toBe("A reply.");
    expect(comments["1"].parentId).toBe("0");
    expect(childrenMap["0"]).toEqual(["1"]);
  });
});

describe("upsertComments: comment resolution", () => {
  it("sets w15:done to 1 for resolved comments and propagates to parent", () => {
    const result = upsertComments(
      {},
      [
        { id: 0, author: "Reviewer A", date: "2026-01-01T00:00:00Z", text: "Top-level comment." },
        { id: 1, author: "Reviewer B", date: "2026-01-01T00:00:00Z", text: "Agreed -- resolved.", parentId: 0, done: true },
      ],
      OPTS
    );
    const exes = commentExesIn(result.commentsExtendedXml);
    expect(exes).toHaveLength(2);
    expect(exes[1].getAttribute("w15:done")).toBe("1");
    expect(exes[0].getAttribute("w15:done")).toBe("1");
  });
});
