import fs from "node:fs";
import vm from "node:vm";
import { describe, expect, it } from "vitest";
import { tokenize } from "../src/criticmarkup/grammar.js";
import { strip } from "../src/criticmarkup/strip.js";

const html = fs.readFileSync(
  new URL("../Markdown Diff to CriticMarkup Track Changes.html", import.meta.url),
  "utf8"
);
const script = html.match(/<script>([\s\S]*?)<\/script>/)?.[1];
if (!script) throw new Error("Markdown diff HTML does not contain an inline script.");

const context = {
  console,
  Uint32Array,
  Set,
  Map,
  Error,
  TypeError,
  Math,
  Number,
  String,
  Array,
  Object,
  RegExp,
};
vm.createContext(context);
vm.runInContext(
  script +
    "\n;globalThis.__markdownDiffCompiler = { compileMarkdownDiff, generateCriticMarkup, projectCriticMarkup };",
  context
);

const { compileMarkdownDiff } = context.__markdownDiffCompiler;

function compile(source, edited, options = {}) {
  const result = compileMarkdownDiff(source, edited, options);
  expect(result.originalView).toBe(source);
  expect(result.acceptedView).toBe(edited);
  expect(result.invariants).toEqual({
    originalViewMatches: true,
    acceptedViewMatches: true,
  });

  const grammar = tokenize(result.criticMarkup, {
    skipBefore: result.protectedPrefixLength,
  });
  expect(grammar.ok).toBe(true);
  expect(
    strip(result.criticMarkup, {
      skipBefore: result.protectedPrefixLength,
    })
  ).toBe(source);

  return result;
}

describe("Markdown diff CriticMarkup compiler: exact views", () => {
  it("preserves an identical document including its final newline", () => {
    const source = "# Heading\n\nNo changes.\n";
    const result = compile(source, source, { protectPreface: false });
    expect(result.criticMarkup).toBe(source);
    expect(result.editCount).toBe(0);
  });

  it.each([
    ["LF", "# H\n\nOld text.\n", "# H\n\nNew text.\n"],
    ["CRLF", "# H\r\n\r\nOld text.\r\n", "# H\r\n\r\nNew text.\r\n"],
    ["CR", "# H\r\rOld text.\r", "# H\r\rNew text.\r"],
  ])("preserves %s line endings byte-for-byte", (_label, source, edited) => {
    compile(source, edited, { protectPreface: false });
  });

  it("treats case, quote style, and whitespace as real edits", () => {
    const source = "# H\n\nProgram  “estimate”\n";
    const edited = '# H\n\nprogram "estimate"\n';
    const result = compile(source, edited, { protectPreface: false });
    expect(result.criticMarkup).not.toBe(source);
    expect(result.criticMarkup).toMatch(/\{~~|\{--|\{\+\+/);
  });

  it("preserves accepted-view order when similar paragraphs are reordered and rewritten", () => {
    const source = "# H\n\nAlpha policy applies.\n\nBeta policy applies.\n";
    const edited = "# H\n\nBeta policy applies broadly.\n\nAlpha policy applies narrowly.\n";
    compile(source, edited, { protectPreface: false, threshold: 0.2 });
  });

  it("falls back safely to whole-region delete/add for a low-similarity rewrite", () => {
    const source = "# H\n\nOne short sentence.\n";
    const edited = "# H\n\nCompletely unrelated replacement language.\n";
    const result = compile(source, edited, {
      protectPreface: false,
      threshold: 0.9,
    });
    expect(result.criticMarkup).toContain("{--");
    expect(result.criticMarkup).toContain("{++");
  });
});

describe("Markdown diff CriticMarkup compiler: protected preface and collisions", () => {
  it("preserves literal CriticMarkup examples in an identical protected preface", () => {
    const preface = "<!-- CriticMarkup: {++add++} {--delete--} -->\n";
    const source = preface + "# H\n\nOld.\n";
    const edited = preface + "# H\n\nNew.\n";
    const result = compile(source, edited, { protectPreface: true });
    expect(result.protectedPrefixLength).toBe(preface.length);
    expect(result.criticMarkup.startsWith(preface)).toBe(true);
  });

  it("rejects edits to a protected preface", () => {
    expect(() =>
      compileMarkdownDiff(
        "<!-- source -->\n# H\nText\n",
        "<!-- edited -->\n# H\nText\n",
        { protectPreface: true }
      )
    ).toThrow(/protected preface differs/i);
  });

  it("rejects literal CriticMarkup openers in editable document content", () => {
    expect(() =>
      compileMarkdownDiff(
        "# H\nLiteral {++example++}.\n",
        "# H\nLiteral {++changed++}.\n",
        { protectPreface: false }
      )
    ).toThrow(/literal CriticMarkup opener/i);
  });
});
