import { describe, expect, it } from "vitest";
import { tokenize } from "../src/criticmarkup/grammar.js";
import { strip } from "../src/criticmarkup/strip.js";
import { parseEdits } from "../src/criticmarkup/parse.js";

describe("tokenize: happy path for all 5 v1 constructs", () => {
  it("insertion {++text++}", () => {
    const r = tokenize("a {++new++} b");
    expect(r.ok).toBe(true);
    expect(r.tokens).toEqual([
      { type: "ins", rawStart: 2, rawEnd: 11, strippedStart: 2, strippedEnd: 2, text: "new" },
    ]);
  });

  it("deletion {--text--}", () => {
    const r = tokenize("a {--old--} b");
    expect(r.ok).toBe(true);
    expect(r.tokens).toEqual([
      { type: "del", rawStart: 2, rawEnd: 11, strippedStart: 2, strippedEnd: 5, text: "old" },
    ]);
  });

  it("substitution {~~old~>new~~}", () => {
    const r = tokenize("The rule {~~shall~>must~~} apply.");
    expect(r.ok).toBe(true);
    expect(r.tokens).toEqual([
      { type: "sub", rawStart: 9, rawEnd: 26, strippedStart: 9, strippedEnd: 14, oldText: "shall", newText: "must" },
    ]);
  });

  it("anchored comment {==text==}{>>comment<<}", () => {
    const r = tokenize("{==all carriers==}{>>Confirm scope.<<}");
    expect(r.ok).toBe(true);
    expect(r.tokens).toEqual([
      {
        type: "comment",
        anchored: true,
        rawStart: 0,
        rawEnd: 38,
        strippedStart: 0,
        strippedEnd: 12,
        highlightText: "all carriers",
        commentText: "Confirm scope.",
      },
    ]);
  });

  it("bare point comment {>>comment<<}", () => {
    const r = tokenize("Fine as-is.{>>Double-check this.<<}");
    expect(r.ok).toBe(true);
    expect(r.tokens).toEqual([
      { type: "comment", anchored: false, rawStart: 11, rawEnd: 35, strippedStart: 11, strippedEnd: 11, commentText: "Double-check this." },
    ]);
  });
});

describe("tokenize: running dual-coordinate offsets across multiple tokens", () => {
  it("AAA{--BBB--}CCC{++DDD++}EEE: stripped offsets account for each prior token's strip-length", () => {
    const r = tokenize("AAA{--BBB--}CCC{++DDD++}EEE");
    expect(r.ok).toBe(true);
    const [del, ins] = r.tokens;
    // stripped text so far: "AAA" (3) + "BBB" (del keeps old text, 3) + "CCC" (3) = 9 before the ins
    expect(del.strippedStart).toBe(3);
    expect(del.strippedEnd).toBe(6);
    expect(ins.strippedStart).toBe(9);
    expect(ins.strippedEnd).toBe(9); // ins strips to "", zero-width
    // raw offsets are just positions in the original string
    expect(del.rawStart).toBe(3);
    expect(ins.rawStart).toBe(15);
  });
});

describe("tokenize: grammar violations", () => {
  it("rejects nesting", () => {
    const r = tokenize("{++a{--b--}c++}");
    expect(r.ok).toBe(false);
    expect(r.error.message).toMatch(/not closed before|nest/i);
  });

  it("rejects an unbalanced/unclosed token", () => {
    const r = tokenize("{++never closes");
    expect(r.ok).toBe(false);
  });

  it("rejects an unbalanced substitution missing ~>", () => {
    const r = tokenize("{~~old only~~}");
    expect(r.ok).toBe(false);
  });

  it("rejects a highlight not immediately followed by a comment", () => {
    const r = tokenize("{==all carriers==} {>>Confirm scope.<<}"); // space in between
    expect(r.ok).toBe(false);
    expect(r.error.message).toMatch(/immediately followed/i);
  });

  it("rejects a highlight with no trailing comment at all", () => {
    const r = tokenize("{==all carriers==}");
    expect(r.ok).toBe(false);
  });

  it("does not treat an unrelated brace as a grammar error", () => {
    const r = tokenize("Section {1} applies, see {2} also.");
    expect(r.ok).toBe(true);
    expect(r.tokens).toEqual([]);
  });
});

describe("strip: the exact §4 transform", () => {
  it("{++x++} -> \"\"", () => {
    expect(strip("a {++new++} b")).toBe("a  b");
  });
  it("{--x--} -> x", () => {
    expect(strip("a {--old--} b")).toBe("a old b");
  });
  it("{~~a~>b~~} -> a", () => {
    expect(strip("The rule {~~shall~>must~~} apply.")).toBe("The rule shall apply.");
  });
  it("{>>x<<} -> \"\"", () => {
    expect(strip("Fine as-is.{>>Double-check this.<<}")).toBe("Fine as-is.");
  });
  it("{==x==} -> x (paired with a stripped-to-empty comment)", () => {
    expect(strip("{==all carriers==}{>>Confirm scope.<<}")).toBe("all carriers");
  });
  it("passes non-token text through unchanged", () => {
    expect(strip("Plain text with {no braces} intact.")).toBe("Plain text with {no braces} intact.");
  });
  it("the worked example from spec §6.2", () => {
    const edited =
      "The rule {~~shall~>must~~} apply to all {--air--} carriers. " +
      "{==all carriers==}{>>Confirm scope includes indirect air carriers.<<}";
    expect(strip(edited)).toBe("The rule shall apply to all air carriers. all carriers");
  });
  it("throws on invalid grammar rather than silently mis-stripping", () => {
    expect(() => strip("{++a{--b--}c++}")).toThrow();
  });
});

describe("parseEdits: edit descriptors in stripped (== exportedMarkdown) coordinates", () => {
  it("returns one descriptor per token, in document order, with the right shape", () => {
    const edits = parseEdits("AAA{--BBB--}CCC{++DDD++}EEE");
    expect(edits).toEqual([
      { type: "del", mdStart: 3, mdEnd: 6, oldText: "BBB", rawStart: 3, rawEnd: 12, wholeParagraph: false },
      { type: "ins", mdPos: 9, newText: "DDD", rawStart: 15, rawEnd: 24, wholeParagraph: false },
    ]);
  });

  it("D1: flags an ins/del token as wholeParagraph when it alone fills its line between blank-line block boundaries", () => {
    const edits = parseEdits("First para.\n\n{++A whole new paragraph.++}\n\nThird para.");
    expect(edits).toEqual([
      { type: "ins", mdPos: 13, newText: "A whole new paragraph.", rawStart: 13, rawEnd: 41, wholeParagraph: true },
    ]);
  });

  it("D1: does not flag an ins/del token that shares its line with other text", () => {
    const edits = parseEdits("Intro {++new ++}text on the same line.");
    expect(edits[0].wholeParagraph).toBe(false);
  });

  it("sub carries both old and new text", () => {
    const edits = parseEdits("The rule {~~shall~>must~~} apply.");
    expect(edits).toEqual([
      { type: "sub", mdStart: 9, mdEnd: 14, oldText: "shall", newText: "must", rawStart: 9, rawEnd: 26 },
    ]);
  });

  it("anchored comment carries the highlight span and comment text separately", () => {
    const edits = parseEdits("{==all carriers==}{>>Confirm scope.<<}");
    expect(edits).toEqual([
      { type: "comment", anchored: true, mdStart: 0, mdEnd: 12, highlightText: "all carriers", commentText: "Confirm scope.", rawStart: 0, rawEnd: 38 },
    ]);
  });

  it("bare comment is a zero-width point anchor", () => {
    const edits = parseEdits("Fine as-is.{>>Double-check this.<<}");
    expect(edits).toEqual([
      { type: "comment", anchored: false, mdPos: 11, commentText: "Double-check this.", rawStart: 11, rawEnd: 35 },
    ]);
  });

  it("throws on invalid grammar", () => {
    expect(() => parseEdits("{++never closes")).toThrow();
  });
});
