import { DOMParser } from "@xmldom/xmldom";
import { describe, expect, it } from "vitest";
import { checkFileExtension, checkPreflight, loadDocxFromBytes, baseName, D4_ERROR_MESSAGE, COMMENT_REPLY_MESSAGE } from "../src/ui/load.js";
import { loadDocxBytes } from "./helpers/docx.js";

describe("checkFileExtension: reject non-docx at the door (spec §6.1)", () => {
  it("accepts .docx, case-insensitively", () => {
    expect(checkFileExtension("policy-draft.docx").ok).toBe(true);
    expect(checkFileExtension("POLICY-DRAFT.DOCX").ok).toBe(true);
  });

  it("rejects anything else with a clear, actionable message", () => {
    const result = checkFileExtension("policy-draft.pdf");
    expect(result.ok).toBe(false);
    expect(result.message).toContain("policy-draft.pdf");
    expect(result.message.toLowerCase()).toContain(".docx");
  });
});

describe("baseName", () => {
  it("strips the .docx extension case-insensitively", () => {
    expect(baseName("policy-draft.docx")).toBe("policy-draft");
    expect(baseName("policy-draft.DOCX")).toBe("policy-draft");
  });
});

describe("checkPreflight: M6a flow-parameterized preflight checks", () => {
  it("passes a clean export in run-review", () => {
    expect(checkPreflight({ counts: { ins: 0, del: 0, sub: 0 }, comments: {} }, "run-review").ok).toBe(true);
  });

  it("passes pre-existing insertions/deletions/comments in run-review (annotation fence down)", () => {
    const result = checkPreflight({ counts: { ins: 1, del: 0, sub: 0 }, comments: { c1: { id: "c1", parentId: null } } }, "run-review");
    expect(result.ok).toBe(true);
  });

  it("blocks on clean export in respond-review flow (nothing to respond to)", () => {
    const result = checkPreflight({ counts: { ins: 0, del: 0, sub: 0 }, comments: {} }, "respond-review");
    expect(result.ok).toBe(false);
    expect(result.message).toContain("nothing here to respond to");
  });

  it("passes annotated export in respond-review flow", () => {
    const result = checkPreflight({ counts: { ins: 1, del: 0, sub: 0 }, comments: {} }, "respond-review");
    expect(result.ok).toBe(true);
  });
});

describe("loadDocxFromBytes: end-to-end against real fixtures", () => {
  it("accepts a clean document and tags it with the extension-stripped filename", async () => {
    const bytes = loadDocxBytes("plain-paragraphs");
    const result = await loadDocxFromBytes(bytes, { originalFilename: "plain-paragraphs.docx", DOMParserImpl: DOMParser });
    expect(result.ok).toBe(true);
    expect(result.filename).toBe("plain-paragraphs");
    expect(result.exported.markdown).toContain("Redline export from: plain-paragraphs.docx");
  });

  it("rejects a non-.docx filename before ever touching the bytes", async () => {
    const result = await loadDocxFromBytes(new ArrayBuffer(0), { originalFilename: "notes.txt" });
    expect(result.ok).toBe(false);
    expect(result.message).toContain("notes.txt");
  });

  it("loads a document with pre-existing comments in run-review", async () => {
    const bytes = loadDocxBytes("comments-threaded");
    const result = await loadDocxFromBytes(bytes, { originalFilename: "comments-threaded.docx", DOMParserImpl: DOMParser, flowType: "run-review" });
    expect(result.ok).toBe(true);
  });
});
