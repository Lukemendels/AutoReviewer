import { describe, expect, it } from "vitest";
import { extractCandidates, selectSingleCandidate } from "../src/envelope.js";

const DOC = "line one\nline two\nline three\nline four\nline five\n"; // exportedLength reference

describe("extractCandidates: the common case (one fence)", () => {
  it("auto-selects the single document-sized fenced block, chatter and all", () => {
    const pasted = "Sure, here's the revised document:\n\n```markdown\n" + DOC + "```\n\nLet me know if you need anything else!";
    const { candidates, noFencesFound } = extractCandidates(pasted, { exportedLength: DOC.length });
    expect(noFencesFound).toBe(false);
    expect(candidates.length).toBe(1);
    expect(candidates[0].content).toBe(DOC);
  });

  it("selectSingleCandidate returns null unless there is exactly one candidate", () => {
    const pasted = "```markdown\n" + DOC + "```";
    expect(selectSingleCandidate(pasted, { exportedLength: DOC.length }).content).toBe(DOC);
  });
});

describe("extractCandidates: verbatim extraction, never trim/append", () => {
  it("does not append a newline when the fence's closing ``` immediately follows the last character", () => {
    const noTrailingNewline = "line one\nline two"; // no final \n
    const pasted = "```markdown\n" + noTrailingNewline + "```";
    const { candidates } = extractCandidates(pasted, { exportedLength: noTrailingNewline.length });
    expect(candidates[0].content).toBe(noTrailingNewline);
    expect(candidates[0].content.endsWith("\n")).toBe(false);
  });

  it("does not trim a leading/trailing blank line that is genuinely part of the content", () => {
    const withBlankLines = "\nline one\n\nline two\n";
    const pasted = "```markdown\n" + withBlankLines + "```";
    const { candidates } = extractCandidates(pasted, { exportedLength: withBlankLines.length });
    expect(candidates[0].content).toBe(withBlankLines);
  });
});

describe("extractCandidates: zero fences", () => {
  it("treats the entire paste as the candidate but flags it (not a silent guess among multiple fences -- there simply are none)", () => {
    const pasted = "no fences here, just the raw text " + DOC;
    const { candidates, noFencesFound, fences } = extractCandidates(pasted, { exportedLength: DOC.length });
    expect(noFencesFound).toBe(true);
    expect(fences).toEqual([]);
    expect(candidates).toEqual([{ content: pasted, fenceInfo: null }]);
  });
});

describe("extractCandidates: multiple plausible fences -> never guess", () => {
  it("returns every document-sized fence rather than picking one", () => {
    const pasted = "First attempt:\n```markdown\n" + DOC + "```\n\nActually, here's a better one:\n```markdown\n" + DOC + "x```";
    const { candidates } = extractCandidates(pasted, { exportedLength: DOC.length });
    expect(candidates.length).toBe(2);
    expect(selectSingleCandidate(pasted, { exportedLength: DOC.length })).toBeNull();
  });

  it("presents the last fence in the paste as the first candidate", () => {
    const first = DOC;
    const second = DOC + "more text to keep it document-sized";
    const pasted = "```markdown\n" + first + "```\n\n```markdown\n" + second + "```";
    const { candidates } = extractCandidates(pasted, { exportedLength: DOC.length });
    expect(candidates[0].content).toBe(second);
    expect(candidates[1].content).toBe(first);
  });

  it("filters out excerpt-sized fences (below the 50% document-length threshold)", () => {
    const excerpt = "a short excerpt";
    const pasted = "```markdown\n" + excerpt + "```\n\n```markdown\n" + DOC + "```";
    const { candidates, fences } = extractCandidates(pasted, { exportedLength: DOC.length });
    expect(fences.length).toBe(2);
    expect(candidates.length).toBe(1);
    expect(candidates[0].content).toBe(DOC);
  });
});

describe("extractCandidates: language tag handling", () => {
  it("captures the language tag but doesn't require it", () => {
    const pastedWithLang = "```markdown\n" + DOC + "```";
    const pastedNoLang = "```\n" + DOC + "```";
    expect(extractCandidates(pastedWithLang, { exportedLength: DOC.length }).candidates[0].fenceInfo.lang).toBe("markdown");
    expect(extractCandidates(pastedNoLang, { exportedLength: DOC.length }).candidates[0].fenceInfo.lang).toBe("");
  });
});
