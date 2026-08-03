import { describe, expect, it } from "vitest";
import { compileMarkdownDiff, deriveProtectedPrefixLength } from "../src/markdown-diff.js";
import { strip } from "../src/criticmarkup/strip.js";

describe("deterministic Markdown diff compiler", () => {
  const preface = "<!-- AutoReviewer export -->\n<!-- CriticMarkup: {++add++} {--delete--} -->\n\n";

  it("preserves an unchanged document exactly", () => {
    const source = `${preface}A plain paragraph.\n`;
    const result = compileMarkdownDiff(source, source);
    expect(result.criticMarkup).toBe(source);
    expect(result.editGroups).toBe(0);
  });

  it("derives a valid inline substitution while preserving the original view", () => {
    const source = `${preface}The rule shall apply.\n`;
    const revised = `${preface}The rule must apply.\n`;
    const prefix = deriveProtectedPrefixLength(source);
    const result = compileMarkdownDiff(source, revised, { protectedPrefixLength: prefix });
    expect(result.criticMarkup).toContain("{~~shall~>must~~}");
    expect(strip(result.criticMarkup, { skipBefore: prefix })).toBe(source);
  });

  it("treats case and punctuation changes as real edits", () => {
    const source = `${preface}Program “Estimate”.\n`;
    const revised = `${preface}program \"estimate\".\n`;
    const result = compileMarkdownDiff(source, revised);
    expect(result.criticMarkup).not.toBe(source);
  });

  it("fails closed when the protected exporter preface changes", () => {
    const source = `${preface}Text.\n`;
    const revised = `<!-- altered -->\n\nText.\n`;
    expect(() => compileMarkdownDiff(source, revised)).toThrow(/protected AutoReviewer preface/i);
  });
});
