// Whole-paragraph insert/delete mechanics (spec §9.1 step 7, §4; M3b plan D3/D4/D5).
// Written and red before applyWholeParagraphEdits has a real implementation.
import { describe, expect, it } from "vitest";
import { DOMParser } from "@xmldom/xmldom";
import { loadDocumentXmlDom, exportFixture } from "./helpers/docx.js";
import { validate } from "../src/validate.js";
import { injectEdits, locateParagraph, locateRun } from "../src/ooxml/inject.js";
import { NS, kid, kids, wAttr, parseXml } from "../src/ooxml/parse.js";

function runText(el) {
  let s = "";
  for (const c of el.children) {
    if (c.localName === "t" || c.localName === "delText") s += c.textContent;
  }
  return s;
}
// Collects run text from a paragraph's direct-child runs AND runs nested one level inside
// w:ins/w:del wrappers (the shape whole-paragraph-insert content and whole-paragraph-
// delete-wrapped runs both take) -- in document order.
function paragraphText(p) {
  let s = "";
  for (const c of p.children) {
    if (c.localName === "r") s += runText(c);
    else if (c.localName === "ins" || c.localName === "del") {
      for (const inner of c.children) if (inner.localName === "r") s += runText(inner);
    }
  }
  return s;
}

async function setup(fixtureName) {
  const exported = await exportFixture(fixtureName);
  const { body } = await loadDocumentXmlDom(fixtureName);
  return { exported, body };
}
function inject(body, edits, sourceMap, opts) {
  return injectEdits(body.ownerDocument, edits, sourceMap, opts);
}
const AUTHOR_OPTS = { author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z" };

describe("whole-paragraph insert: plain paragraph between two existing blocks", () => {
  it("inserts a new <w:p> sibling after the anchor, with the paragraph mark flagged inserted", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    const response = exported.markdown.replace(
      "document.\n\nThis is the second",
      "document.\n{++A whole new paragraph.++}\nThis is the second"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits[0].anchor).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "after" });

    inject(body, result.edits, exported.sourceMap, AUTHOR_OPTS);

    const p0 = locateParagraph(body, [0]);
    const newP = p0.nextSibling;
    expect(newP.localName).toBe("p");
    expect(paragraphText(newP)).toBe("A whole new paragraph.");

    const pPr = kid(newP, "pPr");
    expect(pPr, "new paragraph must carry a pPr for the mark flag").toBeTruthy();
    const rPr = kid(pPr, "rPr");
    expect(rPr).toBeTruthy();
    const markIns = kid(rPr, "ins");
    expect(markIns, "paragraph mark itself must be flagged inserted").toBeTruthy();
    expect(wAttr(markIns, "author")).toBe("AutoReviewer — Test");

    const contentIns = [...newP.children].find((c) => c.localName === "ins");
    expect(contentIns, "the run text itself must also be wrapped in w:ins").toBeTruthy();
    expect(runText(kid(contentIns, "r"))).toBe("A whole new paragraph.");

    // The next paragraph (bodyPath [1] before injection) is untouched and still follows.
    expect(newP.nextSibling.localName).toBe("p");
    expect(paragraphText(newP.nextSibling)).toBe("This is the second paragraph, with nothing tracked.");
  });
});

describe("whole-paragraph insert: list item (D5 best-effort list detection)", () => {
  it("strips the bullet prefix from the inserted text and copies numPr from the adjacent list item", async () => {
    const { exported, body } = await setup("headings-and-lists");
    const response = exported.markdown.replace(
      "- First bullet\n\n- Second bullet",
      "- First bullet\n{++- A new bullet++}\n- Second bullet"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);

    inject(body, result.edits, exported.sourceMap, AUTHOR_OPTS);

    const firstBullet = locateParagraph(body, [5]);
    const newP = firstBullet.nextSibling;
    expect(newP.localName).toBe("p");
    // Prefix stripped from the actual run text.
    expect(paragraphText(newP)).toBe("A new bullet");

    const anchorNumPr = kid(kid(firstBullet, "pPr"), "numPr");
    const newNumPr = kid(kid(newP, "pPr"), "numPr");
    expect(newNumPr, "numPr must be copied from the adjacent list item").toBeTruthy();
    expect(wAttr(kid(newNumPr, "numId"), "val")).toBe(wAttr(kid(anchorNumPr, "numId"), "val"));
  });

  it("leaves a bullet-looking prefix as literal text when there is no adjacent list item to copy numPr from", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    const response = exported.markdown.replace(
      "document.\n\nThis is the second",
      "document.\n{++- Looks like a bullet but isn't one++}\nThis is the second"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);

    inject(body, result.edits, exported.sourceMap, AUTHOR_OPTS);

    const p0 = locateParagraph(body, [0]);
    const newP = p0.nextSibling;
    // Prefix NOT stripped -- no adjacent list item, so D5 leaves it as literal text.
    expect(paragraphText(newP)).toBe("- Looks like a bullet but isn't one");
    expect(kid(kid(newP, "pPr") || {}, "numPr")).toBeFalsy();
  });
});

describe("whole-paragraph insert: at document start and document end", () => {
  it("inserts before the first block when the point resolves before blocks[0]", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    const response = exported.markdown.replace(
      "-->\n\nThis is the first",
      "-->\n{++A brand new opening paragraph.++}\nThis is the first"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits[0].anchor).toEqual({ kind: "paragraphBoundary", bodyPath: [0], edge: "before" });

    // Captured BEFORE injection: bodyPath is positional, and inserting a new paragraph
    // before it shifts what bodyPath [0] resolves to afterward.
    const originalP0 = locateParagraph(body, [0]);

    inject(body, result.edits, exported.sourceMap, AUTHOR_OPTS);

    const newP = originalP0.previousSibling;
    expect(newP.localName).toBe("p");
    expect(paragraphText(newP)).toBe("A brand new opening paragraph.");
    expect(newP.nextSibling).toBe(originalP0);
  });

  it("inserts after the last block when the point resolves after the final block", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    const response = exported.markdown.replace(/compare against\.\n$/, "compare against.\n{++A brand new closing paragraph.++}");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits[0].anchor).toEqual({ kind: "paragraphBoundary", bodyPath: [2], edge: "after" });

    inject(body, result.edits, exported.sourceMap, AUTHOR_OPTS);

    const p2 = locateParagraph(body, [2]);
    const newP = p2.nextSibling;
    expect(newP.localName).toBe("p");
    expect(paragraphText(newP)).toBe("A brand new closing paragraph.");
  });
});

describe("whole-paragraph insert: two consecutive inserts at the same anchor", () => {
  it("processes them in left-to-right document order, each landing immediately after the previous one", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    const response = exported.markdown.replace(
      "document.\n\nThis is the second",
      "document.\n{++First new paragraph.++}{++Second new paragraph.++}\nThis is the second"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits).toHaveLength(2);
    expect(result.edits.every((e) => e.anchor.edge === "after" && e.anchor.bodyPath[0] === 0)).toBe(true);

    inject(body, result.edits, exported.sourceMap, AUTHOR_OPTS);

    const p0 = locateParagraph(body, [0]);
    const first = p0.nextSibling;
    const second = first.nextSibling;
    const original1 = second.nextSibling;
    expect(paragraphText(first)).toBe("First new paragraph.");
    expect(paragraphText(second)).toBe("Second new paragraph.");
    expect(paragraphText(original1)).toBe("This is the second paragraph, with nothing tracked.");
  });
});

describe("whole-paragraph delete: plain paragraph", () => {
  it("wraps every run in <w:del> and flags the paragraph mark deleted", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    const target = "This is the second paragraph, with nothing tracked.";
    const response = exported.markdown.replace(target, `{--${target}--}`);
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits[0].anchor).toEqual({ kind: "wholeParagraphDelete", bodyPath: [1] });

    inject(body, result.edits, exported.sourceMap, AUTHOR_OPTS);

    const p1 = locateParagraph(body, [1]);
    const runWrappers = [...p1.children].filter((c) => c.localName === "del");
    expect(runWrappers).toHaveLength(1);
    expect(kid(runWrappers[0], "r")).toBeTruthy();
    expect(runText(kid(runWrappers[0], "r"))).toBe(target);
    expect(kid(kid(runWrappers[0], "r"), "delText")).toBeTruthy();
    expect(kid(kid(runWrappers[0], "r"), "t")).toBeFalsy();

    const rPr = kid(kid(p1, "pPr"), "rPr");
    const markDel = kid(rPr, "del");
    expect(markDel, "paragraph mark itself must be flagged deleted").toBeTruthy();
    expect(wAttr(markDel, "author")).toBe("AutoReviewer — Test");
  });
});

describe("whole-paragraph delete: list item", () => {
  it("wraps the list item's run in <w:del>, leaving numPr on the (still-present) paragraph", async () => {
    const { exported, body } = await setup("headings-and-lists");
    const response = exported.markdown.replace("- Second bullet", "{--- Second bullet--}");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits[0].anchor).toEqual({ kind: "wholeParagraphDelete", bodyPath: [6] });

    inject(body, result.edits, exported.sourceMap, AUTHOR_OPTS);

    const p6 = locateParagraph(body, [6]);
    const delWrapper = [...p6.children].find((c) => c.localName === "del");
    expect(delWrapper).toBeTruthy();
    expect(runText(kid(delWrapper, "r"))).toBe("Second bullet");
    expect(kid(kid(p6, "pPr"), "numPr"), "numPr stays on the paragraph itself").toBeTruthy();
    expect(kid(kid(kid(p6, "pPr"), "rPr"), "del")).toBeTruthy();
  });
});

describe("whole-paragraph delete: the last body paragraph carrying its own pPr/sectPr", () => {
  it("inserts the paragraph-mark rPr before sectPr, not after (schema-correct CT_PPr child order)", () => {
    // Synthetic minimal document: no fixture has a paragraph-level sectPr (python-docx
    // always emits a body-level one for a single-section document), but real Word
    // documents commonly do on the paragraph right before a section break -- spec §16's
    // named risk ("body's trailing sectPr paragraph must never be deleted" -- this test is
    // the one legitimate case where it *is* the whole-paragraph target, and must not
    // corrupt the sectPr).
    const xml =
      `<w:document xmlns:w="${NS.w}"><w:body>` +
      `<w:p><w:pPr><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr>` +
      `<w:r><w:t>Only paragraph.</w:t></w:r></w:p>` +
      `</w:body></w:document>`;
    const docDoc = parseXml(xml, DOMParser);
    const body = [...docDoc.documentElement.children].find((c) => c.localName === "body");

    const sourceMap = {
      blocks: [{ mdStart: 0, mdEnd: 15, kind: "p", bodyPath: [0], runs: [{ mdStart: 0, mdEnd: 15, runIndex: 0, charOffset: 0 }] }],
      synthetic: [],
      locked: [],
    };
    const exportedMarkdown = "Only paragraph.";
    const response = "{--Only paragraph.--}";
    const result = validate({ responseMarkdown: response, exportedMarkdown, sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits[0].anchor).toEqual({ kind: "wholeParagraphDelete", bodyPath: [0] });

    injectEdits(docDoc, result.edits, sourceMap, AUTHOR_OPTS);

    const p = locateParagraph(body, [0]);
    const pPr = kid(p, "pPr");
    const childNames = [...pPr.children].map((c) => c.localName);
    const rPrIdx = childNames.indexOf("rPr");
    const sectPrIdx = childNames.indexOf("sectPr");
    expect(rPrIdx).toBeGreaterThanOrEqual(0);
    expect(sectPrIdx).toBeGreaterThanOrEqual(0);
    expect(rPrIdx, `pPr children were [${childNames}] -- rPr must come before sectPr`).toBeLessThan(sectPrIdx);
    expect(kid(pPr, "rPr").children[0].localName).toBe("del");
  });
});

describe("whole-paragraph delete: D4 scope cut (pre-existing tracked changes)", () => {
  it("throws a clear, user-actionable error instead of corrupting the paragraph", () => {
    // Synthetic paragraph with a pre-existing <w:ins> direct child -- unreachable through
    // the full validate() pipeline in practice (G1's no-nesting rule already rejects a
    // single {--...--} token whose span would have to swallow the paragraph's own
    // rendered {++...++} synthetic markup), so this is exercised directly against
    // injectEdits with a hand-built edit, the way the plan's ruling frames it: a defensive
    // guard inside inject.js itself, not a gate reachable via a normal response.
    const xml =
      `<w:document xmlns:w="${NS.w}"><w:body>` +
      `<w:p><w:r><w:t>Before. </w:t></w:r>` +
      `<w:ins w:id="1" w:author="Reviewer A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>Inserted text.</w:t></w:r></w:ins>` +
      `</w:p>` +
      `</w:body></w:document>`;
    const docDoc = parseXml(xml, DOMParser);

    const fakeEdit = { type: "del", anchor: { kind: "wholeParagraphDelete", bodyPath: [0] } };
    expect(() => injectEdits(docDoc, [fakeEdit], { blocks: [] }, AUTHOR_OPTS)).toThrow(
      /existing tracked changes.*Accept or reject them in Word first/s
    );
  });
});
