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
    expect(result.invariants).toEqual({ originalViewMatches: true, acceptedViewMatches: true });
  });

  it("derives a valid inline substitution while proving both authoritative views", () => {
    const source = `${preface}The rule shall apply.\n`;
    const revised = `${preface}The rule must apply.\n`;
    const prefix = deriveProtectedPrefixLength(source);
    const result = compileMarkdownDiff(source, revised, { protectedPrefixLength: prefix });
    expect(result.criticMarkup).toContain("{~~shall~>must~~}");
    expect(strip(result.criticMarkup, { skipBefore: prefix })).toBe(source);
    expect(result.originalView).toBe(source);
    expect(result.acceptedView).toBe(revised);
    expect(result.invariants).toEqual({ originalViewMatches: true, acceptedViewMatches: true });
  });

  it("treats case and punctuation changes as real edits", () => {
    const source = `${preface}Program “Estimate”.\n`;
    const revised = `${preface}program \"estimate\".\n`;
    const result = compileMarkdownDiff(source, revised);
    expect(result.criticMarkup).not.toBe(source);
    expect(result.acceptedView).toBe(revised);
  });

  it("fails closed when the protected exporter preface changes", () => {
    const source = `${preface}Text.\n`;
    const revised = `<!-- altered -->\n\nText.\n`;
    expect(() => compileMarkdownDiff(source, revised)).toThrow(/protected AutoReviewer preface/i);
  });
});
