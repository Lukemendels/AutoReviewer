// Multiple edits landing on the same run (M3b plan: "the case the right-to-left peel
// algorithm exists for"). Exercises the full validate() -> injectEdits() pipeline (not
// just splitRun in isolation, per tests/inject.runSplitting.test.js) so the per-run task
// ordering, revision-id allocation, and comment-marker placement are all covered
// together, the way a real ratified edit set would actually flow through the tool.
import { describe, expect, it } from "vitest";
import { loadDocumentXmlDom, exportFixture } from "./helpers/docx.js";
import { validate } from "../src/validate.js";
import { injectEdits, locateParagraph } from "../src/ooxml/inject.js";
import { wAttr } from "../src/ooxml/parse.js";

function runText(el) {
  let s = "";
  for (const c of el.children) {
    if (c.localName === "t" || c.localName === "delText") s += c.textContent;
  }
  return s;
}

// Compact, order-preserving description of a paragraph's direct children, for asserting
// exact DOM sequencing after injection.
function describe_(el) {
  const ln = el.localName;
  if (ln === "r") {
    const cr = [...el.children].find((c) => c.localName === "commentReference");
    if (cr) return { tag: "r", commentReference: wAttr(cr, "id") };
    return { tag: "r", text: runText(el) };
  }
  if (ln === "ins") return { tag: "ins", id: wAttr(el, "id"), author: wAttr(el, "author"), kids: [...el.children].map(describe_) };
  if (ln === "del") return { tag: "del", id: wAttr(el, "id"), author: wAttr(el, "author"), kids: [...el.children].map(describe_) };
  if (ln === "commentRangeStart") return { tag: "commentRangeStart", id: wAttr(el, "id") };
  if (ln === "commentRangeEnd") return { tag: "commentRangeEnd", id: wAttr(el, "id") };
  return { tag: ln };
}
function describeParagraph(p) {
  return [...p.children].filter((c) => c.localName !== "pPr").map(describe_);
}

async function setup(fixtureName) {
  const exported = await exportFixture(fixtureName);
  const { body } = await loadDocumentXmlDom(fixtureName);
  return { exported, body };
}

describe("adjacent edits: two insertions in the same run", () => {
  it("both insertions land in the right place, in document order, each with its own revision id", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    let response = exported.markdown;
    response = response.replace("the first paragraph", "the {++truly ++}first paragraph");
    response = response.replace("plain document.", "{++entire ++}plain document.");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits).toHaveLength(2);

    injectEdits(body.ownerDocument, result.edits, exported.sourceMap, { author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z" });

    const p = locateParagraph(body, [0]);
    const shape = describeParagraph(p);
    // plain-paragraphs.docx has no pre-existing tracked changes, so the revision-id
    // counter (plan: "seeded at max(existing w:ins/w:del w:id)+1") starts at 1.
    expect(shape).toEqual([
      { tag: "r", text: "This is the " },
      { tag: "ins", id: "1", author: "AutoReviewer — Test", kids: [{ tag: "r", text: "truly " }] },
      { tag: "r", text: "first paragraph of a " },
      { tag: "ins", id: "2", author: "AutoReviewer — Test", kids: [{ tag: "r", text: "entire " }] },
      { tag: "r", text: "plain document." },
    ]);
  });
});

describe("adjacent edits: a deletion immediately followed by an insertion (shared boundary)", () => {
  it("processes the insertion (rightmost) before peeling the deletion out of the remainder, preserving final document order", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    let response = exported.markdown;
    // Delete "second" and insert "2nd" immediately after it, both anchored at the same
    // boundary point -- the scenario right-to-left peel ordering exists for.
    response = response.replace("the second paragraph", "the {--second--}{++2nd++} paragraph");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits.map((e) => e.type)).toEqual(["del", "ins"]);

    injectEdits(body.ownerDocument, result.edits, exported.sourceMap, { author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z" });

    const p = locateParagraph(body, [1]);
    const shape = describeParagraph(p);
    expect(shape).toEqual([
      { tag: "r", text: "This is the " },
      { tag: "del", id: shape[1].id, author: "AutoReviewer — Test", kids: [{ tag: "r", text: "second" }] },
      { tag: "ins", id: shape[2].id, author: "AutoReviewer — Test", kids: [{ tag: "r", text: "2nd" }] },
      { tag: "r", text: " paragraph, with nothing tracked." },
    ]);
  });

  it("processes an insertion immediately BEFORE a deletion the same way, ending with insert-then-delete in document order", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    let response = exported.markdown;
    response = response.replace("the second paragraph", "the {++2nd ++}{--second --}paragraph");
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits.map((e) => e.type)).toEqual(["ins", "del"]);

    injectEdits(body.ownerDocument, result.edits, exported.sourceMap, { author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z" });

    const p = locateParagraph(body, [1]);
    const shape = describeParagraph(p);
    expect(shape).toEqual([
      { tag: "r", text: "This is the " },
      { tag: "ins", id: shape[1].id, author: "AutoReviewer — Test", kids: [{ tag: "r", text: "2nd " }] },
      { tag: "del", id: shape[2].id, author: "AutoReviewer — Test", kids: [{ tag: "r", text: "second " }] },
      { tag: "r", text: "paragraph, with nothing tracked." },
    ]);
  });
});

describe("adjacent edits: a deletion and a comment boundary in the same run", () => {
  it("places the comment's start/end markers correctly around an independent, non-overlapping deletion in the same run", async () => {
    const { exported, body } = await setup("plain-paragraphs");
    let response = exported.markdown;
    response = response.replace("A third paragraph exists so", "A {--third--}{++3rd++} paragraph exists so");
    response = response.replace(
      "table/list neighbors in later fixtures",
      "{==table/list neighbors in later fixtures==}{>>Verify this framing.<<}"
    );
    const result = validate({ responseMarkdown: response, exportedMarkdown: exported.markdown, sourceMap: exported.sourceMap });
    expect(result.ok, JSON.stringify(result)).toBe(true);
    expect(result.edits.map((e) => e.type)).toEqual(["del", "ins", "comment"]);

    injectEdits(body.ownerDocument, result.edits, exported.sourceMap, { author: "AutoReviewer — Test", date: "2026-01-01T00:00:00Z" });

    const p = locateParagraph(body, [2]);
    const shape = describeParagraph(p);
    const commentId = shape.find((n) => n.tag === "commentRangeStart").id;
    expect(shape).toEqual([
      { tag: "r", text: "A " },
      { tag: "del", id: shape[1].id, author: "AutoReviewer — Test", kids: [{ tag: "r", text: "third" }] },
      { tag: "ins", id: shape[2].id, author: "AutoReviewer — Test", kids: [{ tag: "r", text: "3rd" }] },
      { tag: "r", text: " paragraph exists so " },
      { tag: "commentRangeStart", id: commentId },
      { tag: "r", text: "table/list neighbors in later fixtures" },
      { tag: "commentRangeEnd", id: commentId },
      { tag: "r", commentReference: commentId },
      { tag: "r", text: " have a plain-paragraph baseline to compare against." },
    ]);
  });
});
